import traceback
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import ppt_converter as pp
import ppt_image as ii

app = FastAPI()

# Configure allowed origins
origins = [
    "https://presently-v1-0.vercel.app/",
    "http://localhost:3000",  # React development server
    "http://127.0.0.1:3000"  # Alternate localhost
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Open to all origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"message": "Hello World"}


@app.post("/generate-presentation/")
async def generate_presentation(request_data: dict):
    try:
        ppt_stream = pp.create_presentation(request_data)
        headers = {"Content-Disposition": f"attachment; filename={request_data['topic']}_presentation.pptx"}
        return StreamingResponse(ppt_stream, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers=headers)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/rate-presentation/")
async def rate_presentation(file: UploadFile):
    if not file.filename.lower().endswith('.pptx'):
        raise HTTPException(
            status_code=400, 
            detail="Only .pptx files are supported"
        )
    
    try:
        contents = await file.read()
        return pp.rate_ppt(contents)
    
    except Exception as e:
        print(traceback.format_exc())
        return JSONResponse(
            status_code=500, 
            content={
                "error": str(e),
                "details": traceback.format_exc()
            }
        )

@app.post("/convert-ppt/")
async def convert_ppt(file: UploadFile):
    if not file.filename.lower().endswith('.pptx'):
        raise HTTPException(
            status_code=400, 
            detail="Only .pptx files are supported"
        )
    
    try:
        ppt_contents = await file.read()
        result = ii.convert_ppt_to_images(ppt_contents)
        return JSONResponse(content=result)
    
    except Exception as e:
        print(traceback.format_exc())
        return JSONResponse(
            status_code=500, 
            content={
                "error": str(e),
                "details": traceback.format_exc()
            }
        )

@app.post("/train-analyse")
async def train(ppt_file: UploadFile = File(...), audio_file: UploadFile = File(...), audio_length: float = Form(...)):
    try:
        ppt_contents = await ppt_file.read()
        audio_contents = await audio_file.read()
        result = pp.analyse_training(ppt_contents, audio_contents, audio_length)
        return JSONResponse(content=result)
    
    except Exception as e:
        print(traceback.format_exc())
        return JSONResponse(
            status_code=500, 
            content={
                "error": str(e),
                "details": traceback.format_exc()
            }
        )

@app.post("/generate-quiz/")
async def generate_quiz_questions(file: UploadFile):
    if not file.filename.lower().endswith('.pptx'):
        raise HTTPException(
            status_code=400,
            detail="Only .pptx files are supported"
        )
    
    try:
        contents = await file.read()
        return pp.generate_quiz(contents)
    
    except Exception as e:
        print(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "details": traceback.format_exc()
            }
        )

@app.post("/evaluate-answer/")
async def evaluate_answer(
    audio_file: UploadFile = File(...),
    question: str = Form(...),
    ideal_answer: str = Form(...)
):
    try:
        if not audio_file.content_type.startswith('audio/'):
            raise HTTPException(
                status_code=400,
                detail="File provided is not an audio file"
            )
        
        audio_contents = await audio_file.read()
        
        # First transcribe the audio
        transcribed_answer = pp.transcribe_audio(audio_contents)
        print("Transcribed answer:", transcribed_answer)
        
        # Create a prompt for the LLaMA model to evaluate the answer
        prompt = f"""
        You are an AI assistant evaluating a presenter's answer to a quiz question asked by a audience.
        
        Question: {question}
        Ideal Answer: {ideal_answer}
        Presenter's Answer: {transcribed_answer}
        
        Provide 2-3 constructive suggestions for improvement. Focus on:
        1. Evaluate users answer based on the ideal answer for the question
        2. Accuracy of the content 
        3. Completeness of the answer
        4. Clarity of expression
        
        Format your response as 2-3 clear, concise sentences that are encouraging and helpful.
        Keep each suggestion to one sentence.
        Do not return any additional text.
        """
        
        # Get response from Groq model
        response = pp.client.chat.completions.create(
            model="llama3-70b-8192",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=1024,
            top_p=1,
            stream=False,
        )
        
        suggestions = response.choices[0].message.content.strip()
        
        return JSONResponse(content={
            "transcribed_answer": transcribed_answer,
            "suggestions": suggestions
        })
        
    except Exception as e:
        print(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "details": traceback.format_exc()
            }
        )