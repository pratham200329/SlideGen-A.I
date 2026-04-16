import json
import os
import uuid
from pathlib import Path
from typing import List, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from ai_generator import edit_slide_with_ai, generate_presentation_json, normalize_presentation_for_export
from ppt_generator import build_pptx_file

load_dotenv()

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(
    title="SlideGen AI API",
    description="Backend API for the SlideGen AI presentation generator.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class PresentationRequest(BaseModel):
    topic: str
    description: str = ""
    slides: int = 8
    tone: str = "Professional"
    audience: str = "Business"
    theme: str = "modern"
    presenterName: str = ""


class GeneratedSlide(BaseModel):
    id: Optional[str] = ""
    aiEdited: Optional[bool] = False
    slide_type: str
    title: str
    subtitle: Optional[str] = ""
    bullets: Optional[List[str]] = []
    left_title: Optional[str] = ""
    left_points: Optional[List[str]] = []
    right_title: Optional[str] = ""
    right_points: Optional[List[str]] = []
    milestones: Optional[List[str]] = []
    chart_title: Optional[str] = ""
    chart_type: Optional[str] = ""
    chart_points: Optional[List[dict]] = []
    layout: Optional[str] = ""
    design: Optional[dict] = {}
    animation_hint: Optional[str] = ""


class GeneratedPresentation(BaseModel):
    title: str
    slides: List[GeneratedSlide]
    theme: str
    suggestions: List[str] = []
    downloadUrl: str
    logs: List[str] = []


class SlideEditRequest(BaseModel):
    presentationTitle: str
    theme: str = "modern"
    slideIndex: int
    slide: dict
    instruction: str
    allSlides: Optional[List[dict]] = []


class ExportPresentationRequest(BaseModel):
    title: str
    theme: str = "modern"
    slides: List[dict]


def _attach_slide_metadata(slides: List[dict]) -> List[dict]:
    normalized = []
    for index, slide in enumerate(slides or [], start=1):
        if not isinstance(slide, dict):
            continue
        copy_slide = dict(slide)
        copy_slide["id"] = str(copy_slide.get("id") or f"slide-{index}")
        copy_slide["aiEdited"] = bool(copy_slide.get("aiEdited", False))
        normalized.append(copy_slide)
    return normalized


history = []


@app.post("/generate")
def generate_presentation(request: PresentationRequest, api_request: Request):
    payload = request.dict()
    print("Request received:", payload)
    logs = [
        "Preparing Groq prompt...",
        "Calling Groq API to generate content...",
    ]
    try:
        generated = generate_presentation_json(**payload)
    except ValueError as exc:
        print("Validation error:", exc)
        raise HTTPException(status_code=502, detail=f"Invalid Groq response: {exc}") from exc
    except RuntimeError as exc:
        detail = str(exc)
        if "GROQ_API_KEY is required" in detail:
            detail = "AI backend API key is missing or invalid. Please configure GROQ_API_KEY."
        lowered = detail.lower()
        if any(token in lowered for token in ["503", "429", "unavailable", "timeout", "service unavailable", "busy"]):
            detail = "AI servers are busy right now. Please try again in a moment."
        print("Runtime error handling request:", detail)
        raise HTTPException(status_code=502, detail=detail) from exc
    except Exception as exc:
        print("Unexpected backend error:", exc)
        raise HTTPException(status_code=500, detail="Unexpected backend error: " + str(exc)) from exc

    print("Presentation generation succeeded.")
    logs.append("Groq response received.")
    logs.append("Building PowerPoint file...")

    generated_slides = _attach_slide_metadata(generated.get("slides", []))
    generated["slides"] = generated_slides

    file_id = f"slidegen_{uuid.uuid4().hex}.pptx"
    file_path = OUTPUT_DIR / file_id
    build_pptx_file(generated, file_path)
    print("Presentation file created:", str(file_path))

    logs.append("Presentation file ready.")

    download_url = str(api_request.url_for("download_presentation", file_name=file_id))
    history_entry = {
        "id": file_id,
        "topic": request.topic,
        "tone": request.tone,
        "audience": request.audience,
        "theme": generated.get("theme", request.theme),
        "slides": len(generated["slides"]),
        "downloadUrl": download_url,
    }
    history.insert(0, history_entry)
    if len(history) > 10:
        history.pop()

    return {
        "title": generated["title"],
        "slides": generated["slides"],
        "theme": generated.get("theme", request.theme),
        "suggestions": generated.get("suggestions", []),
        "downloadUrl": download_url,
        "logs": logs,
    }


@app.post("/slides/edit")
def edit_single_slide(request: SlideEditRequest):
    if request.slideIndex < 0:
        raise HTTPException(status_code=400, detail="slideIndex must be >= 0")

    try:
        edited = edit_slide_with_ai(
            presentation_title=request.presentationTitle,
            theme=request.theme,
            slide=request.slide,
            instruction=request.instruction,
            context_slides=request.allSlides or [],
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Unexpected edit error: {exc}") from exc

    edited_slide = dict(edited)
    edited_slide["id"] = str((request.slide or {}).get("id") or f"slide-{request.slideIndex + 1}")
    edited_slide["aiEdited"] = True

    return {
        "slideIndex": request.slideIndex,
        "slide": edited_slide,
    }


@app.post("/export")
def export_presentation(request: ExportPresentationRequest, api_request: Request):
    try:
        normalized = normalize_presentation_for_export(
            title=request.title,
            theme=request.theme,
            slides=request.slides,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid presentation payload: {exc}") from exc

    normalized["slides"] = _attach_slide_metadata(normalized.get("slides", []))

    file_id = f"slidegen_{uuid.uuid4().hex}.pptx"
    file_path = OUTPUT_DIR / file_id
    build_pptx_file(normalized, file_path)

    download_url = str(api_request.url_for("download_presentation", file_name=file_id))
    return {
        "title": normalized["title"],
        "theme": normalized["theme"],
        "slides": normalized["slides"],
        "downloadUrl": download_url,
    }


@app.get("/download/{file_name}")
def download_presentation(file_name: str):
    file_path = OUTPUT_DIR / file_name
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path=file_path, filename=file_name, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")


@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.get("/history")
def get_history():
    return history
