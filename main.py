import asyncio
import csv
import io
import json
import math
import posixpath
import re
import secrets
from collections import deque
from collections.abc import AsyncIterator, Iterator
from contextlib import asynccontextmanager
from enum import Enum
from functools import reduce
from itertools import chain
from pathlib import PurePath, PurePosixPath
from struct import unpack_from
from threading import BoundedSemaphore, Condition
from time import monotonic
from typing import Annotated, cast
from zipfile import (
    ZIP_DEFLATED,
    ZIP_STORED,
    BadZipFile,
    LargeZipFile,
    ZipFile,
    ZipInfo,
)

import pandas as pd
import uvicorn
from defusedxml.common import DefusedXmlException
from defusedxml.ElementTree import ParseError, fromstring, iterparse
from fastapi import FastAPI, HTTPException, Query, Request, Response, status
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.xml import DEFUSEDXML
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import UploadFile
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.formparsers import MultiPartException
from starlette.middleware.body_limit import RequestBodyLimitMiddleware

MAX_UPLOAD_BYTES = 100 * 1024 * 1024
MAX_REQUEST_OVERHEAD_BYTES = 1024 * 1024
MAX_REQUEST_BYTES = MAX_UPLOAD_BYTES + MAX_REQUEST_OVERHEAD_BYTES
MAX_CHARS_LIMIT = 1_000_000
MAX_OVERLAP_LIMIT = 100
MAX_WORKBOOK_SHEETS = 100
MAX_WORKBOOK_CELLS = 2_000_000
MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLUMNS = 16_384
MAX_ARCHIVE_MEMBERS = 4_096
MAX_ARCHIVE_CENTRAL_DIRECTORY_BYTES = 16 * 1024 * 1024
MAX_ARCHIVE_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_ARCHIVE_MEMBER_BYTES = 256 * 1024 * 1024
MAX_ARCHIVE_COMPRESSION_RATIO = 200.0
MIN_COMPRESSION_RATIO_MEMBER_BYTES = 1024 * 1024
MAX_OPC_METADATA_BYTES = 4 * 1024 * 1024
MAX_WORKSHEET_MEMORY_BYTES = 256 * 1024 * 1024
MAX_OOXML_UNCOMPRESSED_BYTES = 256 * 1024 * 1024
MAX_OOXML_ELEMENTS = 6_500_000
MAX_OOXML_DEPTH = 256
MAX_OOXML_TEXT_BYTES = 64 * 1024 * 1024
MAX_OOXML_ATTRIBUTE_BYTES = 64 * 1024 * 1024
MAX_OOXML_ATTRIBUTES = 8_000_000
MAX_STYLESHEET_ELEMENTS = 250_000
MAX_SHARED_STRINGS = MAX_WORKBOOK_CELLS
MAX_OUTPUT_BYTES = 32 * 1024 * 1024
MAX_OUTPUT_CHUNKS = 4_096
MAX_LANDSCAPE_SERIALIZATION_PROBES = 16_384
MAX_UPLOAD_SECONDS = 120.0
MAX_CONCURRENT_TRANSFORMATIONS = 1
MAX_QUEUED_TRANSFORMATIONS = 7
MAX_CONVERSION_QUEUE_SECONDS = 30.0
MAX_CONVERSION_REQUESTS_PER_WINDOW = 30
CONVERSION_RATE_LIMIT_WINDOW_SECONDS = 60.0
UPLOAD_READ_SIZE = 1024 * 1024
SUPPORTED_EXCEL_TYPES = {
    ".xlsx": frozenset(
        (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/octet-stream",
        )
    ),
    ".xlsm": frozenset(
        ("application/vnd.ms-excel.sheet.macroenabled.12", "application/octet-stream")
    ),
    ".xltx": frozenset(
        (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
            "application/octet-stream",
        )
    ),
    ".xltm": frozenset(
        (
            "application/vnd.ms-excel.template.macroenabled.12",
            "application/octet-stream",
        )
    ),
}
REQUIRED_XLSX_MEMBERS = frozenset(
    (
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
    )
)
CELL_REFERENCE_PATTERN = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$")
OOXML_CONTENT_TYPE_ROLES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml": "worksheet",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": "styles",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedstrings+xml": "sharedStrings",
}
OOXML_WORKBOOK_CONTENT_TYPES = frozenset(
    (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        "application/vnd.ms-excel.sheet.macroenabled.main+xml",
        "application/vnd.ms-excel.template.macroenabled.main+xml",
    )
)
OOXML_RELATIONSHIP_ROLES = {
    "/worksheet": "worksheet",
    "/styles": "styles",
    "/sharedstrings": "sharedStrings",
}
OOXML_ROOT_ROLES = {
    "worksheet": "worksheet",
    "styleSheet": "styles",
    "sst": "sharedStrings",
    "workbook": "workbook",
}
WORKBOOK_PARSE_ERRORS = (
    BadZipFile,
    DefusedXmlException,
    EOFError,
    InvalidFileException,
    KeyError,
    LargeZipFile,
    OSError,
    ParseError,
    pd.errors.ParserError,
    UnicodeError,
    ValueError,
)
ERROR_DETAIL_CODES = {
    "Another workbook transformation is already running.": "transformer_busy",
    "Compressed HTTP request bodies are not supported.": "compressed_request_unsupported",
    "Content-Type must be multipart/form-data.": "multipart_required",
    "Encrypted Excel workbooks are not supported.": "encrypted_workbook_unsupported",
    "External workbook part relationships are not supported.": "external_relationship_unsupported",
    "Invalid Content-Length header.": "invalid_content_length",
    "Multi-disk Excel archives are not supported.": "multi_disk_archive_unsupported",
    "Submit exactly one Excel file in the 'file' field.": "single_file_required",
    "The uploaded file contains an invalid package relationship.": "invalid_package_relationship",
    "The uploaded file contains an invalid worksheet reference.": "invalid_worksheet_reference",
    "The uploaded file contains unsafe archive paths.": "unsafe_archive_path",
    "The uploaded file has an invalid archive directory.": "invalid_archive_directory",
    "The uploaded file has an unsupported media type.": "unsupported_file_media_type",
    "The uploaded file has an unsupported workbook package layout.": "unsupported_workbook_layout",
    "The uploaded file is not a readable Excel workbook.": "invalid_workbook",
    "Transformation exceeds the 32MB output limit.": "output_too_large",
    "Transformation exceeds the landscape work limit.": "landscape_work_limit_exceeded",
    "Transformation exceeds the output chunk limit.": "output_chunk_limit_exceeded",
    "Unsupported Excel file type. Use .xlsx, .xlsm, .xltx, or .xltm.": "unsupported_file_extension",
    "Workbook archive directory exceeds the permitted limit.": "archive_directory_too_large",
    "Workbook compression ratio exceeds the permitted limit.": "archive_compression_ratio_too_high",
    "Workbook contains too many archive members.": "too_many_archive_members",
    "Workbook expands beyond the permitted size.": "archive_expands_too_large",
    "Workbook package metadata exceeds the permitted limit.": "package_metadata_too_large",
    "Workbook package roles are inconsistent.": "package_roles_inconsistent",
    "Workbook upload timed out.": "upload_timeout",
    "Workbook uses an unsupported archive compression method.": "unsupported_archive_compression",
    "Workbook XML expands beyond the permitted size.": "xml_too_large",
    "Workbook XML structure exceeds the permitted limit.": "xml_structure_too_large",
    "Workbook XML text exceeds the permitted limit.": "xml_text_too_large",
    "Worksheet requires too much decoded memory.": "worksheet_memory_too_large",
    "ZIP64 Excel archives are not supported.": "zip64_archive_unsupported",
    "Upload must be 100MB or smaller.": "upload_too_large",
    "Too many workbook transformations are waiting. Please retry shortly.": "transformation_queue_full",
    "Workbook transformation queue timed out. Please retry shortly.": "transformation_queue_timeout",
}
ERROR_DETAIL_PATTERNS = (
    ("-cell limit.", "workbook_cell_limit_exceeded"),
    ("-sheet limit.", "workbook_sheet_limit_exceeded"),
    ("exceeds max_chars", "max_chars_too_small"),
)
STATUS_ERROR_CODES = {
    status.HTTP_400_BAD_REQUEST: "bad_request",
    status.HTTP_404_NOT_FOUND: "not_found",
    status.HTTP_408_REQUEST_TIMEOUT: "request_timeout",
    status.HTTP_413_CONTENT_TOO_LARGE: "request_too_large",
    status.HTTP_415_UNSUPPORTED_MEDIA_TYPE: "unsupported_media_type",
    status.HTTP_422_UNPROCESSABLE_CONTENT: "unprocessable_content",
    status.HTTP_429_TOO_MANY_REQUESTS: "rate_limited",
    status.HTTP_500_INTERNAL_SERVER_ERROR: "internal_error",
    status.HTTP_503_SERVICE_UNAVAILABLE: "service_unavailable",
}

app = FastAPI(docs_url="/docs", redoc_url="/redoc", openapi_url="/openapi.json")
app.add_middleware(RequestBodyLimitMiddleware, max_body_size=MAX_REQUEST_BYTES)
CONVERSION_SEMAPHORE = BoundedSemaphore(MAX_CONCURRENT_TRANSFORMATIONS)
CONVERSION_QUEUE: deque[object] = deque()
CONVERSION_QUEUE_CONDITION = Condition()
CONVERSION_RATE_LIMITS: dict[str, deque[float]] = {}

if not DEFUSEDXML:
    raise RuntimeError("openpyxl XML hardening requires defusedxml.")


def error_code(status_code: int, message: str) -> str:
    """_summary_

    Args:
        status_code (int): HTTP response status.
        message (str): user-facing error message.

    Returns:
        str: stable client-facing error code.
    """
    return ERROR_DETAIL_CODES.get(
        message,
        next(
            (code for needle, code in ERROR_DETAIL_PATTERNS if needle in message),
            STATUS_ERROR_CODES.get(status_code, "error"),
        ),
    )


def error_response_payload(status_code: int, detail) -> dict[str, dict[str, str]]:
    """_summary_

    Args:
        status_code (int): HTTP response status.
        detail: FastAPI exception detail payload.

    Returns:
        dict[str, dict[str, str]]: stable API error response.
    """
    if isinstance(detail, dict) and {"code", "message"} <= detail.keys():
        return {
            "error": {"code": str(detail["code"]), "message": str(detail["message"])}
        }
    message = detail if isinstance(detail, str) else "Invalid request."
    return {"error": {"code": error_code(status_code, message), "message": message}}


@app.exception_handler(StarletteHTTPException)
async def api_http_exception_handler(request: Request, exc: StarletteHTTPException):
    return JSONResponse(
        status_code=exc.status_code,
        content=error_response_payload(exc.status_code, exc.detail),
        headers=exc.headers,
    )


@app.exception_handler(RequestValidationError)
async def api_validation_exception_handler(
    request: Request, exc: RequestValidationError
):
    return JSONResponse(
        status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
        content={
            "error": {
                "code": "invalid_request",
                "message": "Request parameters are invalid.",
            }
        },
    )


def suggested_output_filename() -> str:
    """_summary_

    Returns:
        str: collision-resistant JSON download filename.
    """
    return f"excel-conversion-{secrets.token_hex(8)}.json"


def enforce_conversion_rate_limit(request: Request) -> None:
    """_summary_

    Args:
        request (Request): inbound conversion request.
    """
    now, key = monotonic(), request.client.host if request.client else "unknown"
    bucket = CONVERSION_RATE_LIMITS.setdefault(key, deque())
    cutoff = now - CONVERSION_RATE_LIMIT_WINDOW_SECONDS
    while bucket and bucket[0] <= cutoff:
        bucket.popleft()
    if len(bucket) >= MAX_CONVERSION_REQUESTS_PER_WINDOW:
        retry_after = max(
            1, math.ceil(bucket[0] + CONVERSION_RATE_LIMIT_WINDOW_SECONDS - now)
        )
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail={
                "code": "rate_limited",
                "message": "Too many workbook transformations. Please retry shortly.",
            },
            headers={"Retry-After": str(retry_after)},
        )
    bucket.append(now)


def enqueue_conversion_request(token: object) -> int:
    """_summary_

    Args:
        token (object): request-local queue marker.

    Returns:
        int: one-based queue position.
    """
    with CONVERSION_QUEUE_CONDITION:
        if len(CONVERSION_QUEUE) >= MAX_QUEUED_TRANSFORMATIONS:
            raise HTTPException(
                status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
                detail="Too many workbook transformations are waiting. Please retry shortly.",
                headers={"Retry-After": "1"},
            )
        CONVERSION_QUEUE.append(token)
        CONVERSION_QUEUE_CONDITION.notify_all()
        return len(CONVERSION_QUEUE)


def remove_conversion_request(token: object) -> bool:
    """_summary_

    Args:
        token (object): request-local queue marker.

    Returns:
        bool: whether a queued marker was removed.
    """
    with CONVERSION_QUEUE_CONDITION:
        try:
            CONVERSION_QUEUE.remove(token)
        except ValueError:
            return False
        CONVERSION_QUEUE_CONDITION.notify_all()
        return True


def notify_conversion_queue() -> None:
    """_summary_"""
    with CONVERSION_QUEUE_CONDITION:
        CONVERSION_QUEUE_CONDITION.notify_all()


def wait_for_conversion_turn(token: object, deadline: float) -> bool:
    """_summary_

    Args:
        token (object): request-local queue marker.
        deadline (float): monotonic timestamp after which waiting stops.

    Returns:
        bool: whether the caller acquired the conversion semaphore.
    """
    with CONVERSION_QUEUE_CONDITION:
        while True:
            if token not in CONVERSION_QUEUE:
                return False
            if (
                CONVERSION_QUEUE
                and CONVERSION_QUEUE[0] is token
                and CONVERSION_SEMAPHORE.acquire(blocking=False)
            ):
                CONVERSION_QUEUE.popleft()
                CONVERSION_QUEUE_CONDITION.notify_all()
                return True
            if (remaining := deadline - monotonic()) <= 0:
                try:
                    CONVERSION_QUEUE.remove(token)
                except ValueError:
                    pass
                CONVERSION_QUEUE_CONDITION.notify_all()
                return False
            CONVERSION_QUEUE_CONDITION.wait(min(remaining, 0.5))


@asynccontextmanager
async def conversion_slot() -> AsyncIterator[None]:
    """_summary_

    Raises:
        HTTPException: when the bounded queue is full or waiting expires.

    Returns:
        AsyncIterator[None]: held conversion admission slot.
    """
    token, acquired = object(), False
    enqueue_conversion_request(token)
    try:
        acquired = await run_in_threadpool(
            wait_for_conversion_turn,
            token,
            monotonic() + MAX_CONVERSION_QUEUE_SECONDS,
        )
        if not acquired:
            raise HTTPException(
                status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
                detail="Workbook transformation queue timed out. Please retry shortly.",
                headers={"Retry-After": "1"},
            )
        yield
    finally:
        if not acquired:
            remove_conversion_request(token)
        else:
            CONVERSION_SEMAPHORE.release()
            notify_conversion_queue()


# Define the dropdown choices for FastAPI Docs
class OrientationEnum(str, Enum):
    portrait = "portrait"
    landscape = "landscape"


class OutputBoundedStringIO(io.StringIO):
    def __init__(self, character_limit: int):
        super().__init__()
        self.character_limit = character_limit

    def write(self, value: str) -> int:
        if self.tell() + len(value) > self.character_limit:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
                detail="Transformation exceeds the 32MB output limit.",
            )
        return super().write(value)


def dataframe_to_pipe(
    df: pd.DataFrame,
    *,
    header: bool = True,
    character_limit: int | None = None,
) -> str:
    """_summary_

    Args:
        df (pd.DataFrame): source dataframe to serialize.
        header (bool): whether to serialize dataframe column names.
        character_limit (int | None): maximum in-memory serialized characters.

    Returns:
        str: pipe-delimited dataframe text.
    """
    buffer = (
        io.StringIO()
        if character_limit is None
        else OutputBoundedStringIO(character_limit)
    )
    df.to_csv(buffer, sep="|", index=False, header=header)
    return buffer.getvalue()


def build_landscape_chunk(
    df: pd.DataFrame, start_col_idx: int, end_col_idx: int
) -> str:
    """_summary_

    Args:
        df (pd.DataFrame): source worksheet dataframe.
        start_col_idx (int): inclusive first non-sticky column index.
        end_col_idx (int): exclusive final non-sticky column index.

    Returns:
        str: pipe-delimited chunk with the sticky first column.
    """
    return dataframe_to_pipe(
        pd.concat([df.iloc[:, [0]], df.iloc[:, start_col_idx:end_col_idx]], axis=1)
    )


def validate_excel_upload_headers(request: Request) -> None:
    """_summary_

    Args:
        request (Request): inbound multipart request.
    """
    if (
        request.headers.get("content-encoding", "identity").strip().lower()
        != "identity"
    ):
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail="Compressed HTTP request bodies are not supported.",
        )
    if (
        request.headers.get("content-type", "").partition(";")[0].strip().lower()
        != "multipart/form-data"
    ):
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail="Content-Type must be multipart/form-data.",
        )

    if content_length := request.headers.get("content-length"):
        try:
            if (declared_length := int(content_length)) < 0:
                raise ValueError
            if declared_length > MAX_REQUEST_BYTES:
                raise HTTPException(
                    status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                    detail="Upload must be 100MB or smaller.",
                )
        except ValueError as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Invalid Content-Length header.",
            ) from exc


async def read_single_excel_upload(request: Request) -> bytes:
    """_summary_

    Args:
        request (Request): inbound multipart request.

    Returns:
        bytes: bounded Excel upload contents.
    """
    validate_excel_upload_headers(request)

    try:
        form_context = request.form(
            max_files=1, max_fields=0, max_part_size=MAX_UPLOAD_BYTES
        )
    except (MultiPartException, ValueError) as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Submit exactly one Excel file in the 'file' field.",
        ) from exc

    try:
        async with form_context as form:
            uploads = [
                (key, value)
                for key, value in form.multi_items()
                if isinstance(value, UploadFile)
            ]
            if len(uploads) != 1 or uploads[0][0] != "file":
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Submit exactly one Excel file in the 'file' field.",
                )

            upload = uploads[0][1]
            if (
                suffix := PurePath(upload.filename or "").suffix.lower()
            ) not in SUPPORTED_EXCEL_TYPES:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Unsupported Excel file type. Use .xlsx, .xlsm, .xltx, or .xltm.",
                )
            if (upload.content_type or "").lower() not in SUPPORTED_EXCEL_TYPES[suffix]:
                raise HTTPException(
                    status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
                    detail="The uploaded file has an unsupported media type.",
                )
            if upload.size is not None and upload.size > MAX_UPLOAD_BYTES:
                raise HTTPException(
                    status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                    detail="Upload must be 100MB or smaller.",
                )

            size, payload = 0, bytearray()
            while chunk := await upload.read(UPLOAD_READ_SIZE):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise HTTPException(
                        status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                        detail="Upload must be 100MB or smaller.",
                    )
                payload.extend(chunk)
    except MultiPartException as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Submit exactly one Excel file in the 'file' field.",
        ) from exc

    return bytes(payload)


def validate_zip_directory(contents: bytes) -> None:
    """_summary_

    Args:
        contents (bytes): ZIP-compatible workbook payload.
    """
    signature, lower_bound, cursor = (
        b"PK\x05\x06",
        max(0, len(contents) - 65_557),
        len(contents),
    )
    while (offset := contents.rfind(signature, lower_bound, cursor)) >= 0:
        if len(contents) - offset >= 22:
            (
                _,
                disk_number,
                directory_disk,
                disk_entries,
                total_entries,
                directory_size,
                directory_offset,
                comment_size,
            ) = unpack_from("<4s4H2LH", contents, offset)
            if offset + 22 + comment_size == len(contents):
                if disk_number or directory_disk or disk_entries != total_entries:
                    raise HTTPException(
                        status_code=status.HTTP_400_BAD_REQUEST,
                        detail="Multi-disk Excel archives are not supported.",
                    )
                if (
                    total_entries == 0xFFFF
                    or directory_size == 0xFFFFFFFF
                    or directory_offset == 0xFFFFFFFF
                ):
                    raise HTTPException(
                        status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                        detail="ZIP64 Excel archives are not supported.",
                    )
                if (
                    total_entries > MAX_ARCHIVE_MEMBERS
                    or directory_size > MAX_ARCHIVE_CENTRAL_DIRECTORY_BYTES
                ):
                    raise HTTPException(
                        status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                        detail="Workbook archive directory exceeds the permitted limit.",
                    )
                if directory_offset + directory_size != offset:
                    raise HTTPException(
                        status_code=status.HTTP_400_BAD_REQUEST,
                        detail="The uploaded file has an invalid archive directory.",
                    )
                record_offset, record_count = directory_offset, 0
                while record_offset < offset:
                    if (
                        offset - record_offset < 46
                        or contents[record_offset : record_offset + 4] != b"PK\x01\x02"
                    ):
                        raise HTTPException(
                            status_code=status.HTTP_400_BAD_REQUEST,
                            detail="The uploaded file has an invalid archive directory.",
                        )
                    name_size, extra_size, member_comment_size = unpack_from(
                        "<3H", contents, record_offset + 28
                    )
                    record_offset += 46 + name_size + extra_size + member_comment_size
                    record_count += 1
                    if record_count > MAX_ARCHIVE_MEMBERS:
                        raise HTTPException(
                            status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                            detail="Workbook contains too many archive members.",
                        )
                    if record_offset > offset:
                        raise HTTPException(
                            status_code=status.HTTP_400_BAD_REQUEST,
                            detail="The uploaded file has an invalid archive directory.",
                        )
                if record_count != total_entries:
                    raise HTTPException(
                        status_code=status.HTTP_400_BAD_REQUEST,
                        detail="The uploaded file has an invalid archive directory.",
                    )
                return
        cursor = offset
    raise HTTPException(
        status_code=status.HTTP_400_BAD_REQUEST,
        detail="The uploaded file is not a readable Excel workbook.",
    )


def parse_cell_reference(reference: str) -> tuple[int, int]:
    """_summary_

    Args:
        reference (str): OOXML cell reference.

    Returns:
        tuple[int, int]: one-based row and column coordinates.
    """
    if not (match := CELL_REFERENCE_PATTERN.fullmatch(reference)):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file contains an invalid worksheet reference.",
        )
    letters, row_text = match.groups()
    row, column = (
        int(row_text),
        reduce(
            lambda value, character: value * 26 + ord(character) - 64,
            letters.upper(),
            0,
        ),
    )
    if row > MAX_EXCEL_ROWS or column > MAX_EXCEL_COLUMNS:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file contains an invalid worksheet reference.",
        )
    return row, column


def parse_dimension_reference(reference: str) -> tuple[int, int]:
    """_summary_

    Args:
        reference (str): OOXML worksheet dimension reference.

    Returns:
        tuple[int, int]: maximum one-based row and column coordinates.
    """
    references = reference.split(":")
    if len(references) not in (1, 2):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file contains an invalid worksheet reference.",
        )
    coordinates = tuple(map(parse_cell_reference, references))
    return max(row for row, _ in coordinates), max(column for _, column in coordinates)


def normalize_opc_part_name(part_name: str, base: str = "") -> str:
    """_summary_

    Args:
        part_name (str): OPC part name or relationship target.
        base (str): relationship source directory.

    Returns:
        str: normalized package-relative member name.
    """
    if (
        not part_name
        or "\\" in part_name
        or "://" in part_name
        or "?" in part_name
        or "#" in part_name
    ):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file contains an invalid package relationship.",
        )
    normalized = (
        part_name.lstrip("/")
        if part_name.startswith("/")
        else posixpath.normpath(posixpath.join(base, part_name))
    )
    path, canonical = PurePosixPath(normalized), PurePosixPath(normalized).as_posix()
    if (
        not normalized
        or normalized == "."
        or normalized != canonical
        or path.is_absolute()
        or ".." in path.parts
    ):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file contains an invalid package relationship.",
        )
    return canonical


def is_safe_archive_member_name(name: str) -> bool:
    """_summary_

    Args:
        name (str): ZIP member name.

    Returns:
        bool: whether the member is a canonical relative package path.
    """
    path = PurePosixPath(name)
    return bool(
        name
        and "\\" not in name
        and not path.is_absolute()
        and ".." not in path.parts
        and name == path.as_posix()
    )


def read_opc_metadata(archive: ZipFile, members: dict[str, ZipInfo], member_name: str):
    """_summary_

    Args:
        archive (ZipFile): validated OOXML archive.
        members (dict[str, ZipInfo]): member metadata by name.
        member_name (str): OPC metadata member to parse.

    Returns:
        Element: defused XML root element.
    """
    if (member := members.get(member_name)) is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file is not a readable Excel workbook.",
        )
    if member.file_size > MAX_OPC_METADATA_BYTES:
        raise HTTPException(
            status_code=status.HTTP_413_CONTENT_TOO_LARGE,
            detail="Workbook package metadata exceeds the permitted limit.",
        )
    return fromstring(archive.read(member))


def resolve_ooxml_parts(
    archive: ZipFile, members: tuple[ZipInfo, ...]
) -> tuple[frozenset[str], dict[str, str]]:
    """_summary_

    Args:
        archive (ZipFile): validated OOXML archive.
        members (tuple[ZipInfo, ...]): bounded non-directory archive members.

    Returns:
        tuple[frozenset[str], dict[str, str]]: XML members and semantic part roles.
    """
    member_map = {member.filename: member for member in members}
    content_types = read_opc_metadata(archive, member_map, "[Content_Types].xml")
    xml_extensions, xml_members, roles, workbook_parts = set(), set(), {}, set()

    for element in content_types.iter():
        tag = element.tag.rpartition("}")[-1] if isinstance(element.tag, str) else ""
        content_type = (element.get("ContentType") or "").lower()
        if tag == "Default" and content_type.endswith(("+xml", "/xml")):
            if extension := element.get("Extension"):
                xml_extensions.add(extension.lower().lstrip("."))
        elif tag == "Override" and (part_name := element.get("PartName")):
            normalized_name = normalize_opc_part_name(part_name)
            if content_type.endswith(("+xml", "/xml")):
                xml_members.add(normalized_name)
            if (
                role := OOXML_CONTENT_TYPE_ROLES.get(content_type)
            ) and roles.setdefault(normalized_name, role) != role:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Workbook package roles are inconsistent.",
                )
            if content_type in OOXML_WORKBOOK_CONTENT_TYPES:
                workbook_parts.add(normalized_name)
    content_types.clear()

    if workbook_parts != {"xl/workbook.xml"}:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file has an unsupported workbook package layout.",
        )

    xml_members.update(
        member.filename
        for member in members
        if member.filename.rpartition(".")[2].lower() in xml_extensions
    )
    if any(member_name not in member_map for member_name in roles):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Workbook package roles are inconsistent.",
        )
    relationships = read_opc_metadata(archive, member_map, "xl/_rels/workbook.xml.rels")
    for element in relationships.iter():
        if (
            not isinstance(element.tag, str)
            or element.tag.rpartition("}")[-1] != "Relationship"
        ):
            continue
        relationship_type = (element.get("Type") or "").lower()
        role = next(
            (
                candidate_role
                for suffix, candidate_role in OOXML_RELATIONSHIP_ROLES.items()
                if relationship_type.endswith(suffix)
            ),
            None,
        )
        if role is None:
            continue
        if (element.get("TargetMode") or "").lower() == "external":
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="External workbook part relationships are not supported.",
            )
        target = normalize_opc_part_name(element.get("Target") or "", "xl")
        if target not in member_map or roles.setdefault(target, role) != role:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Workbook package roles are inconsistent.",
            )
        xml_members.add(target)
    relationships.clear()

    return frozenset(xml_members), roles


def prescan_ooxml(
    archive: ZipFile,
    members: tuple[ZipInfo, ...],
    xml_members: frozenset[str],
    roles: dict[str, str],
) -> None:
    """_summary_

    Args:
        archive (ZipFile): validated OOXML archive.
        members (tuple[ZipInfo, ...]): bounded non-directory archive members.
        xml_members (frozenset[str]): metadata-resolved XML member names.
        roles (dict[str, str]): semantic role by member name.
    """
    if (
        sum(member.file_size for member in members if member.filename in xml_members)
        > MAX_OOXML_UNCOMPRESSED_BYTES
    ):
        raise HTTPException(
            status_code=status.HTTP_413_CONTENT_TOO_LARGE,
            detail="Workbook XML expands beyond the permitted size.",
        )

    element_count = text_bytes = attribute_bytes = attribute_count = 0
    stylesheet_elements = shared_strings = sheet_count = 0
    populated_cells = worksheet_shape_cells = 0

    for member in members:
        if member.filename not in xml_members:
            continue

        part_role = roles.get(member.filename)
        current_row = current_column = max_row = max_column = 0
        element_stack = []

        with archive.open(member) as source:
            for event, element in iterparse(source, events=("start", "end")):
                tag = (
                    element.tag.rpartition("}")[-1]
                    if isinstance(element.tag, str)
                    else ""
                )
                if event == "start":
                    if not element_stack and (root_role := OOXML_ROOT_ROLES.get(tag)):
                        if part_role is not None and part_role != root_role:
                            raise HTTPException(
                                status_code=status.HTTP_400_BAD_REQUEST,
                                detail="Workbook package roles are inconsistent.",
                            )
                        part_role = root_role
                    element_stack.append(element)
                    element_count += 1
                    attribute_count += len(element.attrib)
                    attribute_bytes += sum(
                        len(name.encode("utf-8")) + len(value.encode("utf-8"))
                        for name, value in element.attrib.items()
                    )
                    stylesheet_elements += int(part_role == "styles")
                    shared_strings += int(part_role == "sharedStrings" and tag == "si")
                    sheet_count += int(part_role == "workbook" and tag == "sheet")
                    if (
                        len(element_stack) > MAX_OOXML_DEPTH
                        or element_count > MAX_OOXML_ELEMENTS
                        or attribute_count > MAX_OOXML_ATTRIBUTES
                        or attribute_bytes > MAX_OOXML_ATTRIBUTE_BYTES
                        or stylesheet_elements > MAX_STYLESHEET_ELEMENTS
                        or shared_strings > MAX_SHARED_STRINGS
                        or sheet_count > MAX_WORKBOOK_SHEETS
                    ):
                        raise HTTPException(
                            status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                            detail="Workbook XML structure exceeds the permitted limit.",
                        )

                    if (
                        part_role == "worksheet"
                        and tag == "dimension"
                        and (dimension := element.get("ref"))
                    ):
                        dimension_row, dimension_column = parse_dimension_reference(
                            dimension
                        )
                        max_row, max_column = (
                            max(max_row, dimension_row),
                            max(max_column, dimension_column),
                        )
                    elif part_role == "worksheet" and tag == "row":
                        if declared_row := element.get("r"):
                            if (
                                not declared_row.isdecimal()
                                or not 0
                                < (current_row := int(declared_row))
                                <= MAX_EXCEL_ROWS
                            ):
                                raise HTTPException(
                                    status_code=status.HTTP_400_BAD_REQUEST,
                                    detail="The uploaded file contains an invalid worksheet reference.",
                                )
                        else:
                            current_row += 1
                        current_column, max_row, max_column = (
                            0,
                            max(max_row, current_row),
                            max(max_column, 1),
                        )
                    elif part_role == "worksheet" and tag == "c":
                        if cell_reference := element.get("r"):
                            cell_row, cell_column = parse_cell_reference(cell_reference)
                        else:
                            cell_row, cell_column = (
                                max(1, current_row),
                                current_column + 1,
                            )
                        current_row, current_column = cell_row, cell_column
                        max_row, max_column = (
                            max(max_row, cell_row),
                            max(max_column, cell_column),
                        )
                        populated_cells += 1

                    if (
                        populated_cells > MAX_WORKBOOK_CELLS
                        or worksheet_shape_cells + max_row * max_column
                        > MAX_WORKBOOK_CELLS
                    ):
                        raise HTTPException(
                            status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                            detail=f"Workbook exceeds the {MAX_WORKBOOK_CELLS}-cell limit.",
                        )
                    continue

                text_bytes += len((element.text or "").encode("utf-8")) + len(
                    (element.tail or "").encode("utf-8")
                )
                if text_bytes > MAX_OOXML_TEXT_BYTES:
                    raise HTTPException(
                        status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                        detail="Workbook XML text exceeds the permitted limit.",
                    )
                element.clear()
                if len(element_stack) > 1:
                    element_stack[-2].remove(element)
                element_stack.pop()

        worksheet_shape_cells += max_row * max_column


def validate_xlsx_archive(contents: bytes) -> None:
    """_summary_

    Args:
        contents (bytes): Excel upload contents.
    """
    validate_zip_directory(contents)
    try:
        with ZipFile(io.BytesIO(contents)) as archive:
            members = tuple(archive.infolist())
            files = tuple(member for member in members if not member.is_dir())
            names = tuple(member.filename for member in files)
            total_compressed, total_uncompressed = (
                sum(member.compress_size for member in files),
                sum(member.file_size for member in files),
            )

            if len(members) > MAX_ARCHIVE_MEMBERS:
                raise HTTPException(
                    status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                    detail="Workbook contains too many archive members.",
                )
            if len(names) != len(set(names)) or not REQUIRED_XLSX_MEMBERS.issubset(
                names
            ):
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="The uploaded file is not a readable Excel workbook.",
                )
            if any(not is_safe_archive_member_name(name) for name in names):
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="The uploaded file contains unsafe archive paths.",
                )
            if any(member.flag_bits & 0x1 for member in files):
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Encrypted Excel workbooks are not supported.",
                )
            if any(
                member.compress_type not in (ZIP_STORED, ZIP_DEFLATED)
                for member in files
            ):
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Workbook uses an unsupported archive compression method.",
                )
            if total_uncompressed > MAX_ARCHIVE_UNCOMPRESSED_BYTES or any(
                member.file_size > MAX_ARCHIVE_MEMBER_BYTES for member in files
            ):
                raise HTTPException(
                    status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                    detail="Workbook expands beyond the permitted size.",
                )
            if total_uncompressed > MAX_ARCHIVE_COMPRESSION_RATIO * max(
                1, total_compressed
            ) or any(
                member.file_size >= MIN_COMPRESSION_RATIO_MEMBER_BYTES
                and member.file_size
                > MAX_ARCHIVE_COMPRESSION_RATIO * max(1, member.compress_size)
                for member in files
            ):
                raise HTTPException(
                    status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                    detail="Workbook compression ratio exceeds the permitted limit.",
                )
            xml_members, roles = resolve_ooxml_parts(archive, files)
            prescan_ooxml(archive, files, xml_members, roles)
    except HTTPException:
        raise
    except WORKBOOK_PARSE_ERRORS as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file is not a readable Excel workbook.",
        ) from exc


def validate_workbook_shape(workbook: Workbook) -> None:
    """_summary_

    Args:
        workbook (Workbook): opened read-only workbook.
    """
    sheets = tuple(workbook.worksheets)
    if len(sheets) > MAX_WORKBOOK_SHEETS:
        raise HTTPException(
            status_code=status.HTTP_413_CONTENT_TOO_LARGE,
            detail=f"Workbook exceeds the {MAX_WORKBOOK_SHEETS}-sheet limit.",
        )
    if (
        sum((sheet.max_row or 0) * (sheet.max_column or 0) for sheet in sheets)
        > MAX_WORKBOOK_CELLS
    ):
        raise HTTPException(
            status_code=status.HTTP_413_CONTENT_TOO_LARGE,
            detail=f"Workbook exceeds the {MAX_WORKBOOK_CELLS}-cell limit.",
        )


def next_chunk_start(start: int, consumed: int, overlap: int) -> int:
    """_summary_

    Args:
        start (int): current chunk start index.
        consumed (int): number of rows or columns consumed by the chunk.
        overlap (int): requested overlap count.

    Returns:
        int: next chunk start index that always progresses.
    """
    return start + max(1, consumed - min(overlap, max(0, consumed - 1)))


def enforce_chunk_size(
    sheet_name: str, chunk_text: str, max_chars: int, element: str
) -> str:
    """_summary_

    Args:
        sheet_name (str): worksheet name for diagnostics.
        chunk_text (str): serialized candidate chunk.
        max_chars (int): maximum characters allowed per chunk.
        element (str): atomic worksheet element description.

    Returns:
        str: validated chunk text.
    """
    if len(chunk_text) > max_chars:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail=f"{element} in sheet '{sheet_name}' exceeds max_chars by itself.",
        )
    return chunk_text


def chunk_portrait(
    sheet_name: str, df: pd.DataFrame, max_chars: int, overlap: int
) -> Iterator[str]:
    """_summary_

    Args:
        sheet_name (str): worksheet name for diagnostics.
        df (pd.DataFrame): source worksheet dataframe.
        max_chars (int): maximum characters allowed per chunk.
        overlap (int): overlapping row count.

    Returns:
        Iterator[str]: serialized worksheet chunks.
    """
    header_text = enforce_chunk_size(
        sheet_name, dataframe_to_pipe(df.iloc[:0]), max_chars, "Worksheet header"
    )
    csv.field_size_limit(MAX_OUTPUT_BYTES)
    rows = csv.reader(
        io.StringIO(dataframe_to_pipe(df, character_limit=MAX_OUTPUT_BYTES)),
        delimiter="|",
    )
    next(rows, None)

    def new_chunk(seed_rows=()):
        buffer = io.StringIO()
        buffer.write(header_text)
        writer = csv.writer(buffer, delimiter="|", lineterminator="\n")
        tuple(map(writer.writerow, seed_rows))
        return buffer, writer

    buffer, writer = new_chunk()
    tail: deque[list[str]] = deque(maxlen=overlap)
    row_count = 0

    for worksheet_row, row in enumerate(rows, start=2):
        checkpoint = buffer.tell()
        writer.writerow(row)
        if buffer.tell() <= max_chars:
            tail.append(row)
            row_count += 1
            continue

        buffer.seek(checkpoint)
        buffer.truncate()
        if row_count:
            yield buffer.getvalue()
        seed_count = min(overlap, max(0, row_count - 1), len(tail))
        seed_rows = tuple(tail)[-seed_count:] if seed_count else ()

        while True:
            buffer, writer = new_chunk(seed_rows)
            writer.writerow(row)
            if buffer.tell() <= max_chars:
                break
            if seed_rows:
                seed_rows = seed_rows[1:]
                continue
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
                detail=f"Row {worksheet_row} in sheet '{sheet_name}' exceeds max_chars by itself.",
            )

        tail = deque(seed_rows, maxlen=overlap)
        tail.append(row)
        row_count = len(seed_rows) + 1

    if row_count:
        yield buffer.getvalue()


def chunk_landscape(
    sheet_name: str, df: pd.DataFrame, max_chars: int, overlap: int
) -> Iterator[str]:
    """_summary_

    Args:
        sheet_name (str): worksheet name for diagnostics.
        df (pd.DataFrame): source worksheet dataframe.
        max_chars (int): maximum characters allowed per chunk.
        overlap (int): overlapping column count.

    Returns:
        Iterator[str]: serialized worksheet chunks.
    """
    if len(df.columns) <= 1:
        yield enforce_chunk_size(
            sheet_name, dataframe_to_pipe(df), max_chars, "Worksheet"
        )
        return

    start_col_idx, total_cols, serialization_probes = 1, len(df.columns), 0

    def serialize(end_col_idx: int) -> str:
        nonlocal serialization_probes
        serialization_probes += 1
        if serialization_probes > MAX_LANDSCAPE_SERIALIZATION_PROBES:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
                detail="Transformation exceeds the landscape work limit.",
            )
        return build_landscape_chunk(df, start_col_idx, end_col_idx)

    while start_col_idx < total_cols:
        current_col_idx, chunk_text = start_col_idx + 1, serialize(start_col_idx + 1)
        if len(chunk_text) > max_chars:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
                detail=f"Column '{df.columns[start_col_idx]}' in sheet '{sheet_name}' exceeds max_chars with the sticky first column.",
            )

        failed_col_idx, probe_width = None, 2
        while current_col_idx < total_cols:
            candidate_col_idx = min(total_cols, start_col_idx + probe_width)
            candidate_text = serialize(candidate_col_idx)
            if len(candidate_text) > max_chars:
                failed_col_idx = candidate_col_idx
                break
            current_col_idx, chunk_text = candidate_col_idx, candidate_text
            probe_width <<= 1

        if failed_col_idx is not None:
            low_col_idx, high_col_idx = current_col_idx + 1, failed_col_idx - 1
            while low_col_idx <= high_col_idx:
                candidate_col_idx = (low_col_idx + high_col_idx) >> 1
                candidate_text = serialize(candidate_col_idx)
                if len(candidate_text) <= max_chars:
                    current_col_idx, chunk_text, low_col_idx = (
                        candidate_col_idx,
                        candidate_text,
                        candidate_col_idx + 1,
                    )
                else:
                    high_col_idx = candidate_col_idx - 1

        yield chunk_text
        if current_col_idx >= total_cols:
            break
        start_col_idx = next_chunk_start(
            start_col_idx, current_col_idx - start_col_idx, overlap
        )


def json_string_bytes(value: str) -> int:
    """_summary_

    Args:
        value (str): JSON string value.

    Returns:
        int: exact compact JSON UTF-8 byte length.
    """
    return len(
        json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    )


def put_unique(
    output_json: dict[str, str], key: str, value: str, output_bytes: int
) -> int:
    """_summary_

    Args:
        output_json (dict[str, str]): mutable response object.
        key (str): preferred output key.
        value (str): serialized chunk.
        output_bytes (int): current compact JSON response byte count.

    Returns:
        int: updated compact JSON response byte count.
    """
    if len(output_json) >= MAX_OUTPUT_CHUNKS:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Transformation exceeds the output chunk limit.",
        )

    unique_key = (
        key
        if key not in output_json
        else next(
            candidate
            for index in range(2, len(output_json) + 3)
            if (candidate := f"{key} ({index})") not in output_json
        )
    )
    next_output_bytes = (
        output_bytes
        + int(bool(output_json))
        + json_string_bytes(unique_key)
        + 1
        + json_string_bytes(value)
    )
    if next_output_bytes > MAX_OUTPUT_BYTES:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Transformation exceeds the 32MB output limit.",
        )
    output_json[unique_key] = value
    return next_output_bytes


def convert_workbook(
    contents: bytes, orientation: OrientationEnum, max_chars: int, overlap: int
) -> bytes:
    """_summary_

    Args:
        contents (bytes): validated bounded Excel upload.
        orientation (OrientationEnum): worksheet chunking orientation.
        max_chars (int): maximum characters allowed per chunk.
        overlap (int): requested row or column overlap.

    Returns:
        bytes: compact UTF-8 JSON response body.
    """
    validate_xlsx_archive(contents)
    try:
        with pd.ExcelFile(io.BytesIO(contents), engine="openpyxl") as excel_file:
            validate_workbook_shape(excel_file.book)
            output_json, output_bytes, actual_cells = {}, 2, 0

            for sheet_name in excel_file.sheet_names:
                df = cast(
                    pd.DataFrame,
                    excel_file.parse(
                        sheet_name=sheet_name,
                        dtype=object,
                        keep_default_na=False,
                        na_filter=False,
                    ),
                )
                actual_cells += (len(df) + int(bool(len(df.columns)))) * len(df.columns)
                if actual_cells > MAX_WORKBOOK_CELLS:
                    raise HTTPException(
                        status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                        detail=f"Workbook exceeds the {MAX_WORKBOOK_CELLS}-cell limit.",
                    )
                if (
                    int(df.memory_usage(index=True, deep=True).sum())
                    > MAX_WORKSHEET_MEMORY_BYTES
                ):
                    raise HTTPException(
                        status_code=status.HTTP_413_CONTENT_TOO_LARGE,
                        detail="Worksheet requires too much decoded memory.",
                    )

                if df.empty:
                    output_bytes = put_unique(
                        output_json,
                        sheet_name,
                        enforce_chunk_size(
                            sheet_name,
                            dataframe_to_pipe(df),
                            max_chars,
                            "Worksheet header",
                        ),
                        output_bytes,
                    )
                    continue

                chunks = iter(
                    (
                        chunk_portrait
                        if orientation == OrientationEnum.portrait
                        else chunk_landscape
                    )(sheet_name, df, max_chars, overlap)
                )
                first_chunk, second_chunk = next(chunks), next(chunks, None)
                if second_chunk is None:
                    output_bytes = put_unique(
                        output_json, sheet_name, first_chunk, output_bytes
                    )
                else:
                    for index, chunk_text in enumerate(
                        chain((first_chunk, second_chunk), chunks), start=1
                    ):
                        output_bytes = put_unique(
                            output_json,
                            f"{sheet_name} - {index}",
                            chunk_text,
                            output_bytes,
                        )

            response_body = json.dumps(
                output_json, ensure_ascii=False, allow_nan=False, separators=(",", ":")
            ).encode("utf-8")
            if len(response_body) > MAX_OUTPUT_BYTES:
                raise HTTPException(
                    status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
                    detail="Transformation exceeds the 32MB output limit.",
                )
            return response_body
    except HTTPException:
        raise
    except WORKBOOK_PARSE_ERRORS as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The uploaded file is not a readable Excel workbook.",
        ) from exc


@app.middleware("http")
async def add_security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers.setdefault("Cache-Control", "no-store")
    response.headers.setdefault("Referrer-Policy", "no-referrer")
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    return response


@app.get("/health", include_in_schema=False)
async def health_check() -> dict[str, str]:
    return {"status": "ok"}


@app.post(
    "/convert-excel/",
    openapi_extra={
        "requestBody": {
            "required": True,
            "content": {
                "multipart/form-data": {
                    "schema": {
                        "type": "object",
                        "required": ["file"],
                        "properties": {"file": {"type": "string", "format": "binary"}},
                    }
                }
            },
        }
    },
)
async def convert_excel_to_pipe_json(
    request: Request,
    orientation: Annotated[
        OrientationEnum,
        Query(
            description="Choose 'portrait' to split by rows or 'landscape' to split by columns"
        ),
    ] = OrientationEnum.portrait,
    max_chars: Annotated[
        int,
        Query(
            description="Maximum character limit per text chunk",
            ge=50,
            le=MAX_CHARS_LIMIT,
        ),
    ] = 1200,
    overlap: Annotated[
        int,
        Query(
            description="Number of overlapping rows/columns to carry over",
            ge=0,
            le=MAX_OVERLAP_LIMIT,
        ),
    ] = 5,
):
    enforce_conversion_rate_limit(request)
    validate_excel_upload_headers(request)
    async with conversion_slot():
        try:
            async with asyncio.timeout(MAX_UPLOAD_SECONDS):
                contents = await read_single_excel_upload(request)
        except TimeoutError as exc:
            raise HTTPException(
                status_code=status.HTTP_408_REQUEST_TIMEOUT,
                detail="Workbook upload timed out.",
            ) from exc
        return Response(
            content=await run_in_threadpool(
                convert_workbook, contents, orientation, max_chars, overlap
            ),
            media_type="application/json",
            headers={
                "Content-Disposition": f'attachment; filename="{suggested_output_filename()}"'
            },
        )


if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="127.0.0.1",
        port=8080,
        limit_concurrency=8,
        limit_max_requests=1000,
        server_header=False,
        timeout_graceful_shutdown=30,
    )
