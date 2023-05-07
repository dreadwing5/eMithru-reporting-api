import os
from datetime import datetime
from interaction_report import ExcelReportGenerator
from email_sender import EmailSender
from attendance_report import AttendanceReportGenerator
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi import FastAPI, HTTPException
from typing import Dict, List
from dotenv import load_dotenv
load_dotenv()


new_column_names = {
    'title': 'Title',
    'topic': 'Topic',
    'participants': 'Participants',
    'status': 'Status',
    'createdAt': 'Created At',
    'closedAt': 'Closed At',
    'author': 'Author',
    'description': 'Description'
}

columns_order = ['title', 'topic', 'participants', 'status',
                 'createdAt', 'closedAt', 'author', 'description']


app = FastAPI()


origins = [
    "http://localhost",
    "http://localhost:3000",
    "http://localhost:8000",
    "http://localhost:8080",
    "https://report-generator-api.onrender.com",
    "https://cmrit-mentoring-api.onrender.com",
    "https://cmrit-mentoring-tool-frontend-dreadwing5.vercel.app"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class Data(BaseModel):
    data: Dict[str, List[int]]


# FIXME : remove hardcoded values, figure out way to reuse the email component
@app.post("/generate_excel")
async def generate_excel(data: List[Dict]):
    try:
        report = ExcelReportGenerator(data, "data.xlsx")
        report.reindex_and_rename_columns(columns_order, new_column_names) \
            .apply_datetime_conversion(datetime_columns=['Created At', 'Closed At'], date_format='%Y-%m-%dT%H:%M:%S.%fZ')\
              .create_excel_report()
        sender_email = os.getenv("MAIL_ID")
        sender_password = os.getenv("MAIL_PASS")
        subject = "Monthly Interaction Report"
        body = "Please find the monthly interaction report attached."
        recipients = ["immortalosborn@gmail.com"]
        attachment = "data.xlsx"
        email_sender = EmailSender(
            sender_email, sender_password, subject, body, recipients, attachment)
        email_sender.send_email()
        return FileResponse("data.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="data.xlsx")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generate_attendance_report")
async def generate_attendance_report(data: Dict):
    try:
        report = AttendanceReportGenerator(data, "attendance_report.xlsx")
        report.generate_report()
        sender_email = os.getenv("MAIL_ID")
        sender_password = os.getenv("MAIL_PASS")
        subject = "Monthly Interaction Report"
        body = "Please find the monthly interaction report attached."
        recipients = ["immortalosborn@gmail.com"]
        attachment = "attendance_report.xlsx"
        email_sender = EmailSender(
            sender_email, sender_password, subject, body, recipients, attachment)
        email_sender.send_email()
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
