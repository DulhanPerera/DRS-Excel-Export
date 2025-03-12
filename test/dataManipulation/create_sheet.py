from export_details import export_case_details
from config_loader import get_config, connect_db
import logging


def create_excel_sheet():
    """
    Main function to execute the case details export process.
    """
    try:
        logging.info("Starting case details export process...")
        config = get_config()
        db = connect_db(config)
        incident_id = 78910  # Replace with the actual incident ID you want to export
        export_path = config['EXCEL_EXPORT_PATHS']['WIN_DB']
        collection_name = config['COLLECTIONS']['CASE_DETAIL_COLLECTION_NAME']
        
        export_case_details(db, incident_id, export_path, collection_name)
        
        logging.info("Case details export process completed.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

# if __name__ == "__main__":
#     create_excel_sheet()