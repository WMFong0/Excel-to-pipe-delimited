from fastapi import FastAPI, UploadFile, File, Query
from enum import Enum
import pandas as pd
import io
import uvicorn


app = FastAPI(docs_url="/docs", redoc_url="/redoc", openapi_url="/openapi.json")


# Define the dropdown choices for FastAPI Docs
class OrientationEnum(str, Enum):
    portrait = "portrait"
    landscape = "landscape"

def build_portrait_chunk(row_data_list):
    """Combines a list of rows into a single pipe-delimited string with a top header row."""
    chunk_df = pd.concat(row_data_list)
    buffer = io.StringIO()
    chunk_df.to_csv(buffer, sep="|", index=False)
    return buffer.getvalue()

def build_landscape_chunk(header_col_df, col_data_list):
    """Combines a list of columns with the sticky left header column into a single pipe-delimited string."""
    # Combine the left header column with the subset of data columns side-by-side
    combined_df = pd.concat([header_col_df] + col_data_list, axis=1)
    buffer = io.StringIO()
    combined_df.to_csv(buffer, sep="|", index=False)
    return buffer.getvalue()

@app.post("/convert-excel/")
async def convert_excel_to_pipe_json(
    file: UploadFile = File(...),
    orientation: OrientationEnum = Query(OrientationEnum.portrait, description="Choose 'portrait' to split by rows or 'landscape' to split by columns"),
    max_chars: int = Query(1200, description="Maximum character limit per text chunk", ge=50),
    overlap: int = Query(100, description="Number of overlapping rows/columns to carry over", ge=0)
):
    contents = await file.read()
    excel_sheets = pd.read_excel(io.BytesIO(contents), sheet_name=None)
    output_json = {}

    for sheet_name, df in excel_sheets.items():
        if df.empty:
            buffer = io.StringIO()
            df.to_csv(buffer, sep="|", index=False)
            output_json[sheet_name] = buffer.getvalue()
            continue

        sheet_chunks = []

        # ==========================================
        # PORTRAIT MODE: SPLIT BY ROWS (TOP HEADER)
        # ==========================================
        if orientation == OrientationEnum.portrait:
            total_rows = len(df)
            start_row_idx = 0
            safe_overlap = min(overlap, max_chars // 10) # rough safety guard

            while start_row_idx < total_rows:
                current_chunk_rows = []
                current_row_idx = start_row_idx
                
                while current_row_idx < total_rows:
                    next_row = df.iloc[[current_row_idx]]
                    sim_text = build_portrait_chunk(current_chunk_rows + [next_row])
                    
                    if len(sim_text) > max_chars and len(current_chunk_rows) > 0:
                        break
                    
                    current_chunk_rows.append(next_row)
                    current_row_idx += 1
                    
                    if len(sim_text) > max_chars and len(current_chunk_rows) == 1:
                        current_row_idx += 1
                        break

                chunk_text = build_portrait_chunk(current_chunk_rows)
                sheet_chunks.append(chunk_text)

                if current_row_idx >= total_rows:
                    break

                actual_rows_added = current_row_idx - start_row_idx
                step_size = max(1, actual_rows_added - safe_overlap)
                start_row_idx += step_size

        # ==========================================
        # LANDSCAPE MODE: SPLIT BY COLUMNS (LEFT HEADER)
        # ==========================================
        else:
            total_cols = len(df.columns)
            # Column 0 is the sticky header column
            header_col_df = df.iloc[:, [0]]
            
            # The actual columns to chunk start from index 1 to the end
            start_col_idx = 1 
            safe_overlap = min(overlap, total_cols - 2) if total_cols > 2 else 0

            if total_cols <= 1:
                # If there's only 1 column, it cannot be chunked further horizontally
                sheet_chunks.append(build_landscape_chunk(header_col_df, []))
            else:
                while start_col_idx < total_cols:
                    current_chunk_cols = []
                    current_col_idx = start_col_idx
                    
                    while current_col_idx < total_cols:
                        next_col = df.iloc[:, [current_col_idx]]
                        sim_text = build_landscape_chunk(header_col_df, current_chunk_cols + [next_col])
                        
                        if len(sim_text) > max_chars and len(current_chunk_cols) > 0:
                            break
                        
                        current_chunk_cols.append(next_col)
                        current_col_idx += 1
                        
                        if len(sim_text) > max_chars and len(current_chunk_cols) == 1:
                            current_col_idx += 1
                            break

                    chunk_text = build_landscape_chunk(header_col_df, current_chunk_cols)
                    sheet_chunks.append(chunk_text)

                    if current_col_idx >= total_cols:
                        break

                    actual_cols_added = current_col_idx - start_col_idx
                    step_size = max(1, actual_cols_added - safe_overlap)
                    start_col_idx += step_size

        # ==========================================
        # FINAL DICTIONARY KEY STRUCTURING
        # ==========================================
        if len(sheet_chunks) == 1:
            output_json[sheet_name] = sheet_chunks[0]
        else:
            for index, chunk_text in enumerate(sheet_chunks, start=1):
                output_json[f"{sheet_name} - {index}"] = chunk_text

    return output_json

if __name__ == "__main__":
    uvicorn.run("main:app", host="127.0.0.1", port=8080, reload=True)
