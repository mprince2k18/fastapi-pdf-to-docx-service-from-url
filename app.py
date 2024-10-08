import os
import logging
import requests
from datetime import datetime
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from dotenv import load_dotenv
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

# Load environment variables
load_dotenv()

# Initialize the logger
logging.basicConfig(level=logging.INFO)

# Initialize FastAPI
app = FastAPI()

class PDFToDOCXRequest(BaseModel):
    pdf_url: str

@app.post("/convert-pdf-to-docx")
async def convert_pdf_to_docx(request: PDFToDOCXRequest):
    try:
        # Fetch the PDF from the URL
        response = requests.get(request.pdf_url)
        response.raise_for_status()  # Raise an exception if the request failed
        input_stream = response.content
        logging.info("PDF downloaded successfully from URL.")

        # Create credentials instance
        credentials = ServicePrincipalCredentials(
            client_id=os.getenv('PDF_SERVICES_CLIENT_ID'),
            client_secret=os.getenv('PDF_SERVICES_CLIENT_SECRET')
        )
        logging.info("Credentials created successfully.")

        # Create a PDF Services instance
        pdf_services = PDFServices(credentials=credentials)
        logging.info("PDFServices instance created.")

        # Upload the input file
        input_asset = pdf_services.upload(input_stream=input_stream, mime_type=PDFServicesMediaType.PDF)
        logging.info("Input file uploaded successfully.")

        # Create parameters for the job
        export_pdf_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)

        # Create and submit export job
        export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)
        location = pdf_services.submit(export_pdf_job)
        logging.info("Export job submitted successfully.")

        # Get job result
        pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
        result_asset = pdf_services_response.get_result().get_asset()
        stream_asset = pdf_services.get_content(result_asset)
        logging.info("Job completed successfully. Retrieving content.")

        # Save the output to file
        output_file_path = create_output_file_path()
        with open(output_file_path, "wb") as file:
            file.write(stream_asset.get_input_stream())
        logging.info(f"Successfully exported PDF to DOCX: {output_file_path}")

        # Return the file as a download
        return FileResponse(output_file_path, filename=f"exported_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except (requests.RequestException, Exception) as e:
        logging.exception(f"Exception encountered while executing operation: {e}")
        raise HTTPException(status_code=500, detail="Failed to convert PDF to DOCX.")

def create_output_file_path() -> str:
    now = datetime.now()
    time_stamp = now.strftime("%Y-%m-%dT%H-%M-%S")
    os.makedirs("output/ExportPDFToDOCX", exist_ok=True)
    return f"output/ExportPDFToDOCX/export_{time_stamp}.docx"
