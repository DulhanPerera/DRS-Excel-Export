import os
import datetime
from openpyxl import Workbook
from excel.case_details_sheet import create_case_details_sheet
import logging
import sys

logger = logging.getLogger('excel_data_writer')

def export_case_details(db, incident_id, output_path, collection_name):
    """
    Export case details to an Excel file.
    """
    try:
        logger.info(f"Exporting case details for Incident ID: {incident_id}")
        case_collection = db[collection_name]
        case_data = case_collection.find_one({"incident_id": incident_id})
        
        if not case_data:
            logger.error(f"No case details found for Incident ID: {incident_id}")
            sys.exit(1)
        
        logger.info(f"Case data found: {case_data}")

        # Get the case ID
        case_id = case_data.get("case_id")
        if not case_id:
            logger.error("Case ID not found in the case data.")
            sys.exit(1)

        # Generate a timestamp in the format YYYYMMDD_HHMMSS
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        # Create the output file name
        output_file_name = f"Case_{case_id}_{timestamp}.xlsx"
        output_file_save_location = os.path.join(output_path, output_file_name)

        # Check if the file already exists
        counter = 1
        while os.path.exists(output_file_save_location):
            output_file_name = f"Case_{case_id}_{timestamp}_{counter}.xlsx"
            output_file_save_location = os.path.join(output_path, output_file_name)
            counter += 1

        os.makedirs(os.path.dirname(output_file_save_location), exist_ok=True)
        
        excelworkbook = Workbook()
        ws = create_case_details_sheet(excelworkbook, case_data, db)
        
        try:
            excelworkbook.save(output_file_save_location)
            logger.info(f"Case details exported to {output_file_save_location}")
        except Exception as e:
            logger.error(f"Failed to save Excel file: {e}")
            sys.exit(1)
    except Exception as e:
        logger.error(f"Failed to export case details: {e}")
        sys.exit(1)