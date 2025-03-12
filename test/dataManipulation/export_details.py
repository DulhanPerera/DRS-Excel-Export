from create_case_details_sheet import create_case_details_sheet
import os
from openpyxl import Workbook
import logging


def export_case_details(db, incident_id, output_path, collection_name):
    """
    Export case details to an Excel file.
    """
    try:
        logging.info(f"Exporting case details for Incident ID: {incident_id}")
        case_collection = db[collection_name]
        case_data = case_collection.find_one({"incident_id": incident_id})
        
        if not case_data:
            logging.error(f"No case details found for Incident ID: {incident_id}")
            sys.exit(1)
        
        logging.info(f"Case data found: {case_data}")
        
        wb = Workbook()
        ws = create_case_details_sheet(wb, case_data, db)
        
        if not output_path.endswith('.xlsx'):
            output_path = os.path.join(output_path, 'case_details.xlsx')
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        try:
            wb.save(output_path)
            logging.info(f"Case details exported to {output_path}")
        except Exception as e:
            logging.error(f"Failed to save Excel file: {e}")
            sys.exit(1)
    except Exception as e:
        logging.error(f"Failed to export case details: {e}")
        sys.exit(1)