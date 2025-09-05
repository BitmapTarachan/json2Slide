# Pythonベースイメージ
FROM python:3.11-slim

# 作業ディレクトリ
WORKDIR /app

# 依存ライブラリをインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリのソースをコピー
COPY . .
COPY images/ /app/images/

# FastAPIをUvicornで起動
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]