import sys, json, uuid, tempfile, os
from pathlib import Path
from datetime import datetime, timedelta
from fastapi import FastAPI, UploadFile, Query


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

# --- APIモード ---
app = FastAPI()

@app.post("/generate")
async def generate(
    file: UploadFile,
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

# --- CLIモード ---
def cli_main():
    if len(sys.argv) < 3:
        print("Usage: python main.py plan.json out.pptx")
        sys.exit(1)

    plan_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2])
    with plan_path.open("r", encoding="utf-8") as f:
        plan = json.load(f)

    build_pptx_from_plan(plan, out_path)
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
