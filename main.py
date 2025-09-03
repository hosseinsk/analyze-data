from fastapi import FastAPI, UploadFile
import pandas as pd
from io import BytesIO

app = FastAPI()

@app.post("/analyze")
async def analyze_excel(file: UploadFile):
    # خواندن باینری اکسل
    contents = await file.read()
    df = pd.read_excel(BytesIO(contents))
    
    # پردازش نمونه: جمع ستون amount
    if "amount" in df.columns:
        total = df["amount"].sum()
    else:
        total = None

    return {"total_amount": total, "rows": len(df)}
