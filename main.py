import sys, json, uuid, tempfile, os
from pathlib import Path
from datetime import datetime, timedelta
from fastapi import FastAPI, UploadFile, Body, File, Query
from fastapi.responses import JSONResponse
import uvicorn


# --- PPTX生成関数 ---
from json2Slide import build_pptx_from_plan  

# --- APIモード用: Blob設定 ---
BLOB_CONN_STR = os.environ.get("AZURE_STORAGE_CONNECTION_STRING")
CONTAINER_NAME = "pptx-output"

ENV = os.getenv("APP_ENV", "dev")

if BLOB_CONN_STR:
    from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions
    blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONN_STR)
    container_client = blob_service_client.get_container_client(CONTAINER_NAME)
    try:
        container_client.create_container()
    except Exception:
        pass

    def upload_to_blob(output_path: str) -> str:
        """BlobにアップロードしてSAS URLを返す"""
        blob_name = f"{uuid.uuid4()}.pptx"
        blob_client = container_client.get_blob_client(blob_name)
        with open(output_path, "rb") as data:
            blob_client.upload_blob(data, overwrite=True)

        sas_token = generate_blob_sas(
            account_name=blob_service_client.account_name,
            container_name=CONTAINER_NAME,
            blob_name=blob_name,
            account_key=blob_service_client.credential.account_key,
            permission=BlobSasPermissions(read=True),
            expiry=datetime.utcnow() + timedelta(days=7),
        )
        return f"{blob_client.url}?{sas_token}"

app = FastAPI()

# --- APIモード: ファイルアップロード ---
@app.post("/generate")
async def generate(
    file: UploadFile = File(...),
    theme: str = Query(..., description="スライドテーマ（必須）")
):
    # JSONを一時ファイルに保存
    with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tmp:
        tmp.write(await file.read())
        json_path = Path(tmp.name)

    out_path = Path(tempfile.gettempdir()) / f"{uuid.uuid4()}.pptx"

    with json_path.open("r", encoding="utf-8") as f:
        plan = json.load(f)

    build_pptx_from_plan(plan, out_path, themename=theme)

    if not BLOB_CONN_STR:
        if ENV == "dev":
            return {"local_path": str(out_path), "note": "Blob未設定なのでローカル保存"}
        else:
            raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING が未設定です！")

    sas_url = upload_to_blob(out_path)
    return {"url": sas_url}


# --- APIモード: JSON直受け ---
@app.post("/generate-json")
async def generate_json(
    body: str = Body(...),
    theme: str = Query(..., description="スライドテーマ（必須）")
):
    try:
        # JSONパース（壊れていたらINVALID_JSON）
        try:
            plan = json.loads(body)
        except json.JSONDecodeError:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": {
                        "code": "INVALID_JSON",
                        "message": "アップロードされたJSONが壊れています"
                    }
                }
            )

        out_path = Path(tempfile.gettempdir()) / f"{uuid.uuid4()}.pptx"

        # PPTX生成
        try:
            build_pptx_from_plan(plan, out_path, themename=theme)
        except Exception as e:
            return JSONResponse(
                status_code=500,
                content={
                    "success": False,
                    "error": {
                        "code": "BUILD_FAILED",
                        "message": f"PPTX生成に失敗しました: {str(e)}"
                    }
                }
            )

        # Blob保存 or ローカル返却
        if not BLOB_CONN_STR:
            if ENV == "dev":
                return JSONResponse(
                    status_code=200,
                    content={
                        "success": True,
                        "local_path": str(out_path),
                        "note": "Blob未設定なのでローカル保存"
                    }
                )
            else:
                return JSONResponse(
                    status_code=500,
                    content={
                        "success": False,
                        "error": {
                            "code": "NO_BLOB_CONFIG",
                            "message": "AZURE_STORAGE_CONNECTION_STRING が未設定です"
                        }
                    }
                )

        # 正常終了
        sas_url = upload_to_blob(out_path)
        return JSONResponse(status_code=200, content={"success": True, "url": sas_url})

    except Exception as e:
        # 想定外のエラー
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": {
                    "code": "INTERNAL_ERROR",
                    "message": str(e)
                }
            }
        )

# --- CLIモード ---
def cli_main():
    if len(sys.argv) < 4:
        print("Usage: python main.py plan.json out.pptx theme")
        sys.exit(1)

    plan_path = Path(sys.argv[1])
    out_path = sys.argv[2]
    theme = sys.argv[3]

    with plan_path.open("r", encoding="utf-8") as f:
        plan = json.load(f)

    build_pptx_from_plan(plan, out_path, themename=theme)
    print(f"✅ Done: {out_path}")


# --- 実行切替 ---
if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 引数がある → CLIモード
        cli_main()
    else:
        # 引数なし → APIモード
        import uvicorn
        uvicorn.run(app, host="0.0.0.0", port=8000)