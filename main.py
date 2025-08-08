from fastapi import FastAPI, UploadFile, File, Response
from fastapi.middleware.cors import CORSMiddleware
from pymongo import MongoClient
from dotenv import load_dotenv
import os
import pandas as pd
from io import BytesIO
import uvicorn

load_dotenv()  # Load biến từ file .env

MONGO_URI = os.getenv("MONGO_URI")
client = MongoClient(MONGO_URI)
db = client["backend-zma"]
collection = db["users"]  # Tên collection

app = FastAPI()

# Cho phép frontend kết nối
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/import-excel")
async def import_excel(file: UploadFile = File(...)):
    df = pd.read_excel(await file.read())  # Đọc file excel
    data = df.to_dict(orient="records")    # Chuyển thành list dict

    if not data:
        return {"message": "Không có dữ liệu"}

    collection.insert_many(data)  # Lưu tất cả cột luôn
    return {"message": f"Đã nhập {len(data)} dòng"}

@app.get("/export-excel")
def export_excel():
    data = list(collection.find({}, {"_id": 0}))  # Lấy toàn bộ dữ liệu (ẩn _id)
    if not data:
        return {"message": "Không có dữ liệu để xuất"}

    df = pd.DataFrame(data)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    headers = {
        'Content-Disposition': 'attachment; filename="users_export.xlsx"'
    }
    return Response(content=output.read(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

@app.get("/")
def home():
    return {"message": "✅ Backend Python hoạt động!"}

if __name__ == "__main__":
    uvicorn.run("main:app", reload=True)
