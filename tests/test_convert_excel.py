import asyncio
import csv
import io
import re
import unittest
from concurrent.futures import ThreadPoolExecutor
from hashlib import sha256
from struct import pack_into
from time import monotonic
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
from fastapi.testclient import TestClient
from openpyxl.xml import DEFUSEDXML

from main import (
    CONVERSION_QUEUE,
    CONVERSION_QUEUE_CONDITION,
    CONVERSION_RATE_LIMITS,
    CONVERSION_SEMAPHORE,
    MAX_REQUEST_BYTES,
    RequestBodyLimitMiddleware,
    app,
    build_landscape_chunk,
    chunk_landscape,
    chunk_portrait,
    dataframe_to_pipe,
    enqueue_conversion_request,
    notify_conversion_queue,
    remove_conversion_request,
    wait_for_conversion_turn,
)

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def workbook_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """_summary_

    Args:
        sheets (dict[str, pd.DataFrame]): workbook sheets to serialize.

    Returns:
        bytes: in-memory xlsx payload.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tuple(
            df.to_excel(writer, sheet_name=name, index=False)
            for name, df in sheets.items()
        )
    return buffer.getvalue()


def post_workbook(client: TestClient, sheets: dict[str, pd.DataFrame], **params):
    """_summary_

    Args:
        client (TestClient): FastAPI test client.
        sheets (dict[str, pd.DataFrame]): workbook sheets to post.

    Returns:
        Response: HTTP response from the conversion endpoint.
    """
    return client.post(
        "/convert-excel/",
        params=params,
        files={"file": ("input.xlsx", workbook_bytes(sheets), MIME_XLSX)},
    )


def rewrite_workbook(payload: bytes, rewrite) -> bytes:
    """_summary_

    Args:
        payload (bytes): source workbook bytes.
        rewrite (Callable[[str, bytes], bytes]): archive member transformer.

    Returns:
        bytes: rewritten workbook archive.
    """
    source, output = ZipFile(io.BytesIO(payload)), io.BytesIO()
    with source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        tuple(
            target.writestr(
                member.filename, rewrite(member.filename, source.read(member))
            )
            for member in source.infolist()
        )
    return output.getvalue()


def relocate_workbook_member(
    payload: bytes, source_name: str, target_name: str, rewrite
) -> bytes:
    """_summary_

    Args:
        payload (bytes): source workbook bytes.
        source_name (str): existing OPC member name.
        target_name (str): relocated OPC member name.
        rewrite (Callable[[bytes], bytes]): relocated member transformer.

    Returns:
        bytes: workbook with updated content-type and relationship metadata.
    """
    source, output = ZipFile(io.BytesIO(payload)), io.BytesIO()
    source_path, target_path = source_name.encode(), target_name.encode()
    with source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        for member in source.infolist():
            data = source.read(member)
            if member.filename in (
                "[Content_Types].xml",
                "xl/_rels/workbook.xml.rels",
            ):
                data = data.replace(source_path, target_path).replace(
                    source_path.removeprefix(b"xl/"),
                    target_path.removeprefix(b"xl/"),
                )
            target.writestr(
                target_name if member.filename == source_name else member.filename,
                rewrite(data) if member.filename == source_name else data,
            )
    return output.getvalue()


def add_workbook_member(payload: bytes, member_name: str, member_data: bytes) -> bytes:
    """_summary_

    Args:
        payload (bytes): source workbook bytes.
        member_name (str): extra archive member name.
        member_data (bytes): extra archive member content.

    Returns:
        bytes: workbook archive with an additional member.
    """
    source, output = ZipFile(io.BytesIO(payload)), io.BytesIO()
    with source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        tuple(
            target.writestr(member.filename, source.read(member))
            for member in source.infolist()
        )
        target.writestr(member_name, member_data)
    return output.getvalue()


def error_payload(response) -> dict[str, str]:
    """_summary_

    Args:
        response: FastAPI test response.

    Returns:
        dict[str, str]: structured API error payload.
    """
    return response.json()["error"]


class ConvertExcelTests(unittest.TestCase):
    def setUp(self):
        CONVERSION_RATE_LIMITS.clear()
        with CONVERSION_QUEUE_CONDITION:
            CONVERSION_QUEUE.clear()
            CONVERSION_QUEUE_CONDITION.notify_all()
        self.client = TestClient(app)

    def test_portrait_splitting_keeps_every_row(self):
        response = post_workbook(
            self.client,
            {
                "Sheet1": pd.DataFrame(
                    {"id": [1, 2, 3], "value": ["a" * 20, "b" * 20, "c" * 20]}
                )
            },
            max_chars=50,
            overlap=0,
        )

        self.assertEqual(response.status_code, 200)
        text = "\n".join(response.json().values())
        tuple(self.assertIn(value * 20, text) for value in ("a", "b", "c"))

    def test_portrait_serialization_uses_constant_pandas_calls(self):
        dataframe = pd.DataFrame({"id": range(10_000)})

        with patch("main.dataframe_to_pipe", wraps=dataframe_to_pipe) as serializer:
            chunks = tuple(
                chunk_portrait("Rows", dataframe, max_chars=1_000_000, overlap=0)
            )

        self.assertEqual(len(chunks), 1)
        self.assertEqual(serializer.call_count, 2)

    def test_portrait_csv_escaping_is_preserved(self):
        dataframe = pd.DataFrame(
            {"id": [1, 2], "value": ["a|b", 'line 1\nline 2 "quoted"']}
        )
        expected = list(
            csv.reader(io.StringIO(dataframe_to_pipe(dataframe)), delimiter="|")
        )[1:]

        actual = [
            row
            for chunk in chunk_portrait(
                "Escaping", dataframe, max_chars=1_000, overlap=0
            )
            for row in list(csv.reader(io.StringIO(chunk), delimiter="|"))[1:]
        ]

        self.assertEqual(actual, expected)

    def test_landscape_splitting_keeps_every_column(self):
        response = post_workbook(
            self.client,
            {
                "Sheet1": pd.DataFrame(
                    {"id": [1], "a": ["a" * 20], "b": ["b" * 20], "c": ["c" * 20]}
                )
            },
            orientation="landscape",
            max_chars=50,
            overlap=0,
        )

        self.assertEqual(response.status_code, 200)
        text = "\n".join(response.json().values())
        tuple(self.assertIn(value * 20, text) for value in ("a", "b", "c"))

    def test_wide_landscape_serialization_uses_bounded_probes(self):
        dataframe = pd.DataFrame(
            [range(16_384)], columns=[f"column_{index}" for index in range(16_384)]
        )

        with patch(
            "main.build_landscape_chunk", wraps=build_landscape_chunk
        ) as serializer:
            chunks = tuple(
                chunk_landscape("Wide", dataframe, max_chars=1_000_000, overlap=0)
            )

        self.assertEqual(len(chunks), 1)
        self.assertLessEqual(serializer.call_count, 16)

    def test_single_oversized_row_is_rejected(self):
        response = post_workbook(
            self.client,
            {"Sheet1": pd.DataFrame({"id": [1], "value": ["x" * 80]})},
            max_chars=50,
            overlap=0,
        )

        self.assertEqual(response.status_code, 422)
        self.assertEqual(error_payload(response)["code"], "max_chars_too_small")
        self.assertIn("exceeds max_chars", error_payload(response)["message"])

    def test_sheet_name_collision_does_not_overwrite_output(self):
        response = post_workbook(
            self.client,
            {
                "Foo": pd.DataFrame({"id": [1, 2], "value": ["a" * 20, "b" * 20]}),
                "Foo - 1": pd.DataFrame({"id": [3], "value": ["collision"]}),
            },
            max_chars=50,
            overlap=0,
        )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(len(payload), 3)
        self.assertIn("Foo - 1 (2)", payload)
        self.assertIn("collision", payload["Foo - 1 (2)"])

    def test_invalid_workbook_returns_400(self):
        response = self.client.post(
            "/convert-excel/", files={"file": ("input.xlsx", b"not an xlsx", MIME_XLSX)}
        )

        self.assertEqual(response.status_code, 400)

    def test_invalid_workbook_returns_actionable_error_payload(self):
        response = self.client.post(
            "/convert-excel/", files={"file": ("input.xlsx", b"not an xlsx", MIME_XLSX)}
        )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            error_payload(response),
            {
                "code": "invalid_workbook",
                "message": "The uploaded file is not a readable Excel workbook.",
            },
        )

    def test_invalid_query_returns_actionable_error_payload(self):
        response = self.client.post(
            "/convert-excel/",
            params={"max_chars": 1},
            files={
                "file": (
                    "input.xlsx",
                    workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})}),
                    MIME_XLSX,
                )
            },
        )

        self.assertEqual(response.status_code, 422)
        self.assertEqual(error_payload(response)["code"], "invalid_request")

    def test_more_than_one_file_is_rejected(self):
        payload = workbook_bytes({"Sheet1": pd.DataFrame({"id": [1], "value": ["ok"]})})
        response = self.client.post(
            "/convert-excel/",
            files=[
                ("file", ("a.xlsx", payload, MIME_XLSX)),
                ("other", ("b.xlsx", payload, MIME_XLSX)),
            ],
        )

        self.assertEqual(response.status_code, 400)

    def test_excel_na_literal_is_preserved(self):
        response = post_workbook(
            self.client,
            {"Sheet1": pd.DataFrame({"code": ["NA", "NULL", "N/A"]})},
            max_chars=500,
            overlap=0,
        )

        self.assertEqual(response.status_code, 200)
        text = "\n".join(response.json().values())
        tuple(self.assertIn(value, text) for value in ("NA", "NULL", "N/A"))

    def test_raw_request_body_limit_is_installed_before_form_parsing(self):
        middleware = next(
            item
            for item in app.user_middleware
            if item.cls is RequestBodyLimitMiddleware
        )

        self.assertEqual(middleware.kwargs["max_body_size"], MAX_REQUEST_BYTES)

    def test_raw_request_body_limit_counts_streamed_bytes_without_content_length(self):
        async def exercise(limit: int) -> int:
            messages = iter(
                (
                    {"type": "http.request", "body": b"ab", "more_body": True},
                    {"type": "http.request", "body": b"cd", "more_body": False},
                )
            )
            sent = []

            async def downstream(scope, receive, send):
                while (await receive()).get("more_body", False):
                    pass
                await send(
                    {"type": "http.response.start", "status": 204, "headers": []}
                )
                await send({"type": "http.response.body", "body": b""})

            async def receive():
                return next(messages)

            async def send(message):
                sent.append(message)

            await RequestBodyLimitMiddleware(downstream, max_body_size=limit)(
                {
                    "type": "http",
                    "asgi": {"version": "3.0"},
                    "http_version": "1.1",
                    "method": "POST",
                    "scheme": "http",
                    "path": "/",
                    "raw_path": b"/",
                    "query_string": b"",
                    "root_path": "",
                    "headers": [],
                    "client": ("127.0.0.1", 1),
                    "server": ("testserver", 80),
                },
                receive,
                send,
            )
            return next(
                message["status"]
                for message in sent
                if message["type"] == "http.response.start"
            )

        self.assertEqual(asyncio.run(exercise(4)), 204)
        self.assertEqual(asyncio.run(exercise(3)), 413)

    def test_negative_content_length_is_rejected(self):
        payload = workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})})
        response = self.client.post(
            "/convert-excel/",
            headers={"content-length": "-1"},
            files={"file": ("input.xlsx", payload, MIME_XLSX)},
        )

        self.assertEqual(response.status_code, 400)

    def test_non_multipart_request_is_rejected(self):
        response = self.client.post(
            "/convert-excel/",
            content=b"not multipart",
            headers={"content-type": "application/octet-stream"},
        )

        self.assertEqual(response.status_code, 415)

    def test_invalid_headers_are_rejected_before_transformation_queue(self):
        with patch("main.MAX_QUEUED_TRANSFORMATIONS", 0):
            response = self.client.post(
                "/convert-excel/",
                content=b"not multipart",
                headers={"content-type": "application/octet-stream"},
            )

        self.assertEqual(response.status_code, 415)
        self.assertEqual(error_payload(response)["code"], "multipart_required")

    def test_http_content_encoding_is_rejected(self):
        response = self.client.post(
            "/convert-excel/",
            content=b"compressed",
            headers={
                "content-type": "multipart/form-data; boundary=x",
                "content-encoding": "gzip",
            },
        )

        self.assertEqual(response.status_code, 415)

    def test_wrong_file_media_type_is_rejected(self):
        payload = workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})})
        response = self.client.post(
            "/convert-excel/", files={"file": ("input.xlsx", payload, "text/plain")}
        )

        self.assertEqual(response.status_code, 415)

    def test_spooled_upload_works_in_read_only_container(self):
        payload = workbook_bytes(
            {
                "Sheet1": pd.DataFrame(
                    {
                        "value": [
                            "".join(
                                sha256(f"{index}:{block}".encode()).hexdigest()
                                for block in range(4)
                            )
                            for index in range(10_000)
                        ]
                    }
                )
            }
        )
        self.assertGreater(len(payload), 1024 * 1024)

        response = self.client.post(
            "/convert-excel/",
            params={"max_chars": 1_000_000, "overlap": 0},
            files={"file": ("input.xlsx", payload, MIME_XLSX)},
        )

        self.assertEqual(response.status_code, 200)

    def test_archive_expansion_limit_is_enforced(self):
        with patch("main.MAX_ARCHIVE_UNCOMPRESSED_BYTES", 1):
            response = post_workbook(self.client, {"Sheet1": pd.DataFrame({"id": [1]})})

        self.assertEqual(response.status_code, 413)

    def test_archive_directory_member_limit_is_enforced_before_parsing(self):
        with patch("main.MAX_ARCHIVE_MEMBERS", 1):
            response = post_workbook(self.client, {"Sheet1": pd.DataFrame({"id": [1]})})

        self.assertEqual(response.status_code, 413)

    def test_forged_archive_entry_count_is_rejected_before_parsing(self):
        payload = bytearray(
            workbook_bytes({"Sheet1": pd.DataFrame({"id": [1], "value": ["ok"]})})
        )
        directory_offset = payload.rfind(b"PK\x05\x06")
        self.assertGreaterEqual(directory_offset, 0)
        pack_into("<2H", payload, directory_offset + 8, 1, 1)

        response = self.client.post(
            "/convert-excel/",
            files={"file": ("input.xlsx", bytes(payload), MIME_XLSX)},
        )

        self.assertEqual(response.status_code, 400)

    def test_archive_compression_ratio_limit_is_enforced(self):
        with patch("main.MAX_ARCHIVE_COMPRESSION_RATIO", 1.0):
            response = post_workbook(self.client, {"Sheet1": pd.DataFrame({"id": [1]})})

        self.assertEqual(response.status_code, 413)

    def test_output_byte_limit_is_enforced(self):
        with patch("main.MAX_OUTPUT_BYTES", 10):
            response = post_workbook(
                self.client, {"Sheet1": pd.DataFrame({"id": [1], "value": ["content"]})}
            )

        self.assertEqual(response.status_code, 422)

    def test_output_chunk_limit_is_enforced(self):
        with patch("main.MAX_OUTPUT_CHUNKS", 1):
            response = post_workbook(
                self.client,
                {"Sheet1": pd.DataFrame({"id": [1, 2], "value": ["a" * 20, "b" * 20]})},
                max_chars=50,
                overlap=0,
            )

        self.assertEqual(response.status_code, 422)

    def test_single_column_landscape_respects_max_chars(self):
        response = post_workbook(
            self.client,
            {"Sheet1": pd.DataFrame({"value": ["x" * 80]})},
            orientation="landscape",
            max_chars=50,
            overlap=0,
        )

        self.assertEqual(response.status_code, 422)

    def test_header_only_sheet_respects_max_chars(self):
        response = post_workbook(
            self.client,
            {"Sheet1": pd.DataFrame(columns=["x" * 80])},
            max_chars=50,
            overlap=0,
        )

        self.assertEqual(response.status_code, 422)

    def test_malformed_internal_xml_returns_400(self):
        payload = rewrite_workbook(
            workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})}),
            lambda name, data: b"<broken" if name == "xl/workbook.xml" else data,
        )
        response = self.client.post(
            "/convert-excel/", files={"file": ("input.xlsx", payload, MIME_XLSX)}
        )

        self.assertEqual(response.status_code, 400)

    def test_actual_cell_limit_catches_missing_dimension_metadata(self):
        payload = rewrite_workbook(
            workbook_bytes({"Sheet1": pd.DataFrame({"id": range(20)})}),
            lambda name, data: (
                re.sub(rb"<dimension[^>]*/>", b"", data, count=1)
                if name == "xl/worksheets/sheet1.xml"
                else data
            ),
        )
        with (
            patch("main.MAX_WORKBOOK_CELLS", 10),
            patch(
                "main.pd.ExcelFile",
                side_effect=AssertionError(
                    "workbook materialized before XML preflight"
                ),
            ),
        ):
            response = self.client.post(
                "/convert-excel/", files={"file": ("input.xlsx", payload, MIME_XLSX)}
            )

        self.assertEqual(response.status_code, 413)

    def test_sparse_cell_geometry_rejects_before_workbook_materialization(self):
        payload = rewrite_workbook(
            workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})}),
            lambda name, data: (
                re.sub(rb"<dimension[^>]*/>", b"", data, count=1).replace(
                    b'r="A2"', b'r="C1048576"', 1
                )
                if name == "xl/worksheets/sheet1.xml"
                else data
            ),
        )

        with patch(
            "main.pd.ExcelFile",
            side_effect=AssertionError("workbook materialized before XML preflight"),
        ):
            response = self.client.post(
                "/convert-excel/", files={"file": ("input.xlsx", payload, MIME_XLSX)}
            )

        self.assertEqual(response.status_code, 413)

    def test_relocated_worksheet_role_is_limited_before_materialization(self):
        payload = relocate_workbook_member(
            workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})}),
            "xl/worksheets/sheet1.xml",
            "xl/custom/sheet.part",
            lambda data: re.sub(rb"<dimension[^>]*/>", b"", data, count=1).replace(
                b'r="A2"', b'r="C1048576"', 1
            ),
        )

        with patch(
            "main.pd.ExcelFile",
            side_effect=AssertionError("workbook materialized before XML preflight"),
        ):
            response = self.client.post(
                "/convert-excel/", files={"file": ("input.xlsx", payload, MIME_XLSX)}
            )

        self.assertEqual(response.status_code, 413)

    def test_noncanonical_absolute_relationship_target_is_rejected(self):
        payload = rewrite_workbook(
            relocate_workbook_member(
                workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})}),
                "xl/worksheets/sheet1.xml",
                "xl/custom/sheet.part",
                lambda data: data,
            ),
            lambda name, data: (
                data.replace(
                    b'Target="/xl/custom/sheet.part"',
                    b'Target="/xl//custom/sheet.part"',
                )
                if name == "xl/_rels/workbook.xml.rels"
                else data
            ),
        )

        with patch(
            "main.pd.ExcelFile",
            side_effect=AssertionError("workbook materialized before XML preflight"),
        ):
            response = self.client.post(
                "/convert-excel/", files={"file": ("input.xlsx", payload, MIME_XLSX)}
            )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            error_payload(response)["code"], "invalid_package_relationship"
        )

    def test_noncanonical_archive_member_path_is_rejected(self):
        response = self.client.post(
            "/convert-excel/",
            files={
                "file": (
                    "input.xlsx",
                    add_workbook_member(
                        workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})}),
                        "xl//custom/sheet.part",
                        b"<worksheet/>",
                    ),
                    MIME_XLSX,
                )
            },
        )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(error_payload(response)["code"], "unsafe_archive_path")

    def test_xml_element_limit_rejects_before_workbook_materialization(self):
        payload = workbook_bytes({"Sheet1": pd.DataFrame({"id": [1]})})

        with (
            patch("main.MAX_OOXML_ELEMENTS", 1),
            patch(
                "main.pd.ExcelFile",
                side_effect=AssertionError(
                    "workbook materialized before XML preflight"
                ),
            ),
        ):
            response = self.client.post(
                "/convert-excel/", files={"file": ("input.xlsx", payload, MIME_XLSX)}
            )

        self.assertEqual(response.status_code, 413)

    def test_xml_entity_payload_is_rejected(self):
        payload = rewrite_workbook(
            workbook_bytes({"Sheet1": pd.DataFrame({"value": ["ENTITY_SOURCE"]})}),
            lambda name, data: (
                data.replace(
                    b"?>",
                    b'?><!DOCTYPE worksheet [<!ENTITY injected "ENTITY_EXPANDED">]>',
                    1,
                ).replace(b"ENTITY_SOURCE", b"&injected;")
                if name == "xl/worksheets/sheet1.xml"
                else data
            ),
        )
        response = self.client.post(
            "/convert-excel/", files={"file": ("input.xlsx", payload, MIME_XLSX)}
        )

        self.assertTrue(DEFUSEDXML)
        self.assertEqual(response.status_code, 400)

    def test_security_headers_are_present(self):
        response = self.client.get("/health")

        self.assertEqual(response.headers["cache-control"], "no-store")
        self.assertEqual(response.headers["referrer-policy"], "no-referrer")
        self.assertEqual(response.headers["x-content-type-options"], "nosniff")
        self.assertEqual(response.headers["x-frame-options"], "DENY")

    def test_full_transformation_queue_returns_retryable_error(self):
        with patch("main.MAX_QUEUED_TRANSFORMATIONS", 0):
            response = post_workbook(self.client, {"Sheet1": pd.DataFrame({"id": [1]})})

        self.assertEqual(response.status_code, 503)
        self.assertEqual(response.headers["retry-after"], "1")
        self.assertEqual(error_payload(response)["code"], "transformation_queue_full")

    def test_busy_transformer_queues_then_times_out_with_retry(self):
        self.assertTrue(CONVERSION_SEMAPHORE.acquire(blocking=False))
        try:
            with patch("main.MAX_CONVERSION_QUEUE_SECONDS", 0.0):
                response = post_workbook(
                    self.client, {"Sheet1": pd.DataFrame({"id": [1]})}
                )
        finally:
            CONVERSION_SEMAPHORE.release()
            notify_conversion_queue()

        self.assertEqual(response.status_code, 503)
        self.assertEqual(response.headers["retry-after"], "1")
        self.assertEqual(
            error_payload(response)["code"], "transformation_queue_timeout"
        )

    def test_queued_admission_waits_for_semaphore_release(self):
        token, waiter_acquired, initial_held = object(), False, False
        self.assertTrue(CONVERSION_SEMAPHORE.acquire(blocking=False))
        initial_held = True
        try:
            self.assertEqual(enqueue_conversion_request(token), 1)
            with ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(
                    wait_for_conversion_turn, token, monotonic() + 2.0
                )
                self.assertFalse(future.done())
                CONVERSION_SEMAPHORE.release()
                initial_held = False
                notify_conversion_queue()
                waiter_acquired = future.result(timeout=2.0)
            self.assertTrue(waiter_acquired)
        finally:
            remove_conversion_request(token)
            if waiter_acquired or initial_held:
                CONVERSION_SEMAPHORE.release()
                notify_conversion_queue()

    def test_conversion_rate_limit_returns_retryable_error(self):
        with patch("main.MAX_CONVERSION_REQUESTS_PER_WINDOW", 1):
            first = post_workbook(self.client, {"Sheet1": pd.DataFrame({"id": [1]})})
            second = post_workbook(self.client, {"Sheet1": pd.DataFrame({"id": [1]})})

        self.assertEqual(first.status_code, 200)
        self.assertEqual(second.status_code, 429)
        self.assertGreaterEqual(int(second.headers["retry-after"]), 1)
        self.assertEqual(error_payload(second)["code"], "rate_limited")

    def test_conversion_response_has_unique_safe_download_name(self):
        with patch("main.secrets.token_hex", return_value="abc123def4567890"):
            response = post_workbook(self.client, {"Sheet1": pd.DataFrame({"id": [1]})})

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["content-disposition"],
            'attachment; filename="excel-conversion-abc123def4567890.json"',
        )


if __name__ == "__main__":
    unittest.main()
