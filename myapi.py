import traceback
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import JSONResponse, StreamingResponse

app = FastAPI()



@app.get("/")
async def root():
    return {"message": "Hello World"}

