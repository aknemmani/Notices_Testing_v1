# testing_app/main.py

import os
import shutil
from fastapi.middleware.cors import CORSMiddleware

from fastapi import FastAPI, Request, UploadFile, File, Depends, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from datetime import datetime

from testing_database import (
    init_testing_db,
    get_testing_db,
    TestingNotice,
)
from testing_service import (
    write_generated_row,
    write_generated_row_gpt_5_1,
    write_generated_row_gpt_5_mini,
    is_pdf_in_generated,
    is_pdf_in_master,
    load_and_compare_testing_excel,      # overall
    load_and_compare_gemini_only,
    load_and_compare_gpt_5_1_only,
    load_and_compare_gpt_5_mini_only,
    calculate_overall_accuracy,
    calculate_category_accuracy,
    calculate_disconnect_late_accuracy,
    calculate_notice_date_accuracy,
    calculate_date_accuracy,
    calculate_correct_row_counts
)

from gemini_service import extract_testing_fields_from_pdf  # reuse existing file
from gpt_service import (
    extract_testing_fields_from_pdf_gpt_5_1,
    extract_testing_fields_from_pdf_gpt_5_mini,
)

# Independent testing app
app = FastAPI(title="Notice Testing App")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def on_startup():
    init_testing_db()


# Static + templates
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def testing_root(request: Request):
    return templates.TemplateResponse("testing_index.html", {"request": request})


# ---------- Testing API endpoints ----------


@app.post("/testing/upload-pdf")
async def testing_upload_pdf(
    file: UploadFile = File(...),
    db: Session = Depends(get_testing_db),
):
    """
    Upload single/multiple PDFs for testing.
    Validates that PDF name exists in master sheet first.
    """
    # Check if PDF exists in master sheet
    if not is_pdf_in_master(file.filename):
        raise HTTPException(
            status_code=400,
            detail=f"PDF '{file.filename}' not found in Master sheet. Please add it first.",
        )

    os.makedirs("testing_pdfs", exist_ok=True)
    file_location = os.path.join("testing_pdfs", file.filename)

    with open(file_location, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    notice = TestingNotice(
        pdf_filename=file.filename,
        pdf_path=file_location,
    )
    db.add(notice)
    db.commit()
    db.refresh(notice)

    return {"testing_notice_id": notice.id, "pdf_filename": file.filename}


@app.post("/testing/bulk-process")
async def testing_bulk_process(db: Session = Depends(get_testing_db)):
    """
    Run Gemini testing for testing notices that do NOT yet
    have a row in the generated sheet.
    """
    notices = db.query(TestingNotice).all()

    processed = 0
    skipped = 0
    errors = []

    for notice in notices:
        if not notice.pdf_path or not os.path.exists(notice.pdf_path):
            skipped += 1
            continue

        if is_pdf_in_generated(notice.pdf_filename, model="gemini"):
            skipped += 1
            continue

        try:
            with open(notice.pdf_path, "rb") as f:
                pdf_bytes = f.read()

            testing_result = extract_testing_fields_from_pdf(pdf_bytes)

            excel_fields = {
                "Vendor Account Number": testing_result.get("vendor_account_number", "NA"),
                "Vendor Name": testing_result.get("vendor_name", "NA"),
                "Service Address": testing_result.get("service_address", "NA"),
                "Notice Category": testing_result.get("notice_category", "Others"),
                "Notice Date": testing_result.get("notice_date", "NA"),
                "Impact Date": testing_result.get("impact_date", "NA"),
                "Impact Amount": testing_result.get("impact_amount", "NA"),
            }

            write_generated_row(notice.pdf_filename, excel_fields)
            processed += 1

        except Exception as e:
            errors.append({"testing_notice_id": notice.id, "error": str(e)})

    return {"processed": processed, "skipped": skipped, "errors": errors}


@app.post("/testing/bulk-process-gpt-5-1")
async def testing_bulk_process_gpt_5_1(db: Session = Depends(get_testing_db)):
    """
    Run GPT 5.1 testing for notices that do NOT yet
    have a row in the GPT 5.1 generated sheet.
    """
    notices = db.query(TestingNotice).all()

    processed = 0
    skipped = 0
    errors = []

    for notice in notices:
        if not notice.pdf_path or not os.path.exists(notice.pdf_path):
            skipped += 1
            continue

        if is_pdf_in_generated(notice.pdf_filename, model="gpt-5.1"):
            skipped += 1
            continue

        try:
            with open(notice.pdf_path, "rb") as f:
                pdf_bytes = f.read()

            testing_result = extract_testing_fields_from_pdf_gpt_5_1(pdf_bytes)

            excel_fields = {
                "Vendor Account Number": testing_result.get("vendor_account_number", "NA"),
                "Vendor Name": testing_result.get("vendor_name", "NA"),
                "Service Address": testing_result.get("service_address", "NA"),
                "Notice Category": testing_result.get("notice_category", "Others"),
                "Notice Date": testing_result.get("notice_date", "NA"),
                "Impact Date": testing_result.get("impact_date", "NA"),
                "Impact Amount": testing_result.get("impact_amount", "NA"),
            }

            write_generated_row_gpt_5_1(notice.pdf_filename, excel_fields)
            processed += 1

        except Exception as e:
            errors.append({"testing_notice_id": notice.id, "error": str(e)})

    return {"processed": processed, "skipped": skipped, "errors": errors}


@app.post("/testing/bulk-process-gpt-5-mini")
async def testing_bulk_process_gpt_5_mini(db: Session = Depends(get_testing_db)):
    """
    Run GPT 5-mini testing for notices that do NOT yet
    have a row in the GPT 5-mini generated sheet.
    """
    notices = db.query(TestingNotice).all()

    processed = 0
    skipped = 0
    errors = []

    for notice in notices:
        if not notice.pdf_path or not os.path.exists(notice.pdf_path):
            skipped += 1
            continue

        if is_pdf_in_generated(notice.pdf_filename, model="gpt-5-mini"):
            skipped += 1
            continue

        try:
            with open(notice.pdf_path, "rb") as f:
                pdf_bytes = f.read()

            testing_result = extract_testing_fields_from_pdf_gpt_5_mini(pdf_bytes)

            excel_fields = {
                "Vendor Account Number": testing_result.get("vendor_account_number", "NA"),
                "Vendor Name": testing_result.get("vendor_name", "NA"),
                "Service Address": testing_result.get("service_address", "NA"),
                "Notice Category": testing_result.get("notice_category", "Others"),
                "Notice Date": testing_result.get("notice_date", "NA"),
                "Impact Date": testing_result.get("impact_date", "NA"),
                "Impact Amount": testing_result.get("impact_amount", "NA"),
            }

            write_generated_row_gpt_5_mini(notice.pdf_filename, excel_fields)
            processed += 1

        except Exception as e:
            errors.append({"testing_notice_id": notice.id, "error": str(e)})

    return {"processed": processed, "skipped": skipped, "errors": errors}


@app.get("/testing/comparison-results")
async def testing_comparison_results():
    """
    Read master vs generated from Notices_Testing.xlsx
    and return comparison for the testing UI.
    """
    try:
        return load_and_compare_testing_excel()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/testing/comparison-results-gemini")
async def testing_comparison_results_gemini():
    """
    Master vs Gemini only.
    """
    try:
        return load_and_compare_gemini_only()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/testing/comparison-results-gpt-5-1")
async def testing_comparison_results_gpt_5_1():
    """
    Master vs GPT 5.1 only.
    """
    try:
        return load_and_compare_gpt_5_1_only()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/testing/comparison-results-gpt-5-mini")
async def testing_comparison_results_gpt_5_mini():
    """
    Master vs GPT 5-mini only.
    """
    try:
        return load_and_compare_gpt_5_mini_only()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/testing/analytics/overall-accuracy")
async def get_overall_accuracy():
    """
    Return overall vendor identification accuracy for each model.
    """
    try:
        return calculate_overall_accuracy()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/testing/analytics/category-accuracy")
async def get_category_accuracy():
    """
    Return category-wise mismatch counts for each model.
    """
    try:
        return calculate_category_accuracy()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
@app.get("/testing/analytics/disconnect-late-accuracy")
async def get_disconnect_late_accuracy():
    """
    Return Impact Date and Impact Amount accuracy for Disconnect and Late Notice.
    """
    try:
        return calculate_disconnect_late_accuracy()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
@app.get("/testing/analytics/impact-date")
async def get_disconnect_late_accuracy():
    """
    Return Impact Date and Impact Amount accuracy for Disconnect and Late Notice.
    """
    try:
        return calculate_date_accuracy()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/testing/analytics/notice-date-accuracy")
async def get_notice_date_accuracy():
    """
    Return Notice Date accuracy for all notice types.
    """
    try:
        return calculate_notice_date_accuracy()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
@app.get("/testing/analytics/correct-row-counts")
async def get_correct_row_counts():
    return calculate_correct_row_counts()

    

