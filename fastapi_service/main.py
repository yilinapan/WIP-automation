from __future__ import annotations

import datetime
import urllib.parse
from io import BytesIO
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import Response

import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from processor import process_uploaded_excel, parse_output_filename_from_input  # type: ignore  # noqa: E402


app = FastAPI(
    title="WIP 轉貼紙 API",
    description="使用既有的處理邏輯，提供給 Make / 自動化流程呼叫的轉檔服務。",
)


@app.get("/health")
def health_check() -> dict:
    return {"status": "ok"}


@app.get("/wake")
def wake() -> dict:
    return {"status": "warm", "timestamp": datetime.datetime.utcnow().isoformat()}


@app.post("/convert")
async def convert_wip_to_sticker(
    file: UploadFile = File(..., description="WIP Excel 檔案（.xlsx / .xls）"),
    sheets: Optional[List[str]] = None,
) -> Response:
    """
    上傳一份 WIP Excel，回傳已套版好的貼紙 Excel 檔。

    - `file`: 原始 WIP 檔
    - `sheets`: 可選，多個工作表名稱（query string: ?sheets=Sheet1&sheets=Sheet2）。
      若未提供，則與 Streamlit 相同：預設處理全部工作表。
    """
    if file.filename is None or file.filename == "":
        raise HTTPException(status_code=400, detail="檔名為空，請確認上傳的檔案。")

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="檔案內容為空，請確認上傳的檔案。")

        # Swagger / 呼叫端如果送出空字串（例如 sheets=），會變成 [""]，
        # 這裡把空字串與全空白的值濾掉，避免誤判為「沒有可處理的工作表」。
        if sheets:
            cleaned = [s.strip() for s in sheets if s and s.strip()]
            sheet_names: Optional[List[str]] = cleaned or None
        else:
            sheet_names = None

        excel_bytes, _df = process_uploaded_excel(
            content,
            sheet_names=sheet_names,
            hangtag_image=None,
            bag_image=None,
        )

        # 產生輸出檔名（沿用原本的命名邏輯）
        output_name = parse_output_filename_from_input(file.filename)

        encoded_name = urllib.parse.quote(output_name)
        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
        )
    except HTTPException:
        # 已經是我們主動拋出的錯誤，直接往外丟
        raise
    except Exception as e:
        # 其他未預期錯誤，統一包成 500
        raise HTTPException(status_code=500, detail=f"轉檔失敗：{e}")

