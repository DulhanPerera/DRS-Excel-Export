# main.py
import logging
import logging.config
import sys  # Add this import
from config_and_db import get_config, connect_db
from excel_export import export_case_details

def main():
    """
    Main function to execute the case details export process.
    """
    try:
        logging.config.fileConfig('Config/logger/loggers.ini')
        logger = logging.getLogger('excel_data_writer')
        
        logger.info("Starting case details export process...")
        config = get_config()
        db = connect_db(config)
        incident_id = 78910  # Replace with the actual incident ID you want to export
        export_path = config['EXCEL_EXPORT_PATHS']['WIN_DB']
        collection_name = config['COLLECTIONS']['CASE_DETAIL_COLLECTION_NAME']
        
        export_case_details(db, incident_id, export_path, collection_name)
        
        logger.info("Case details export process completed.")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()