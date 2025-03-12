import logging
import logging.config
from utils.config_loader import get_config
from utils.connectDB import get_db_connection
from excel.excel_writer import export_case_details
import sys

# Load logger configuration
logging.config.fileConfig('config/logger/loggers.ini')
logger = logging.getLogger('excel_data_writer')

def main():
    """
    Main function to execute the case details export process.
    """
    try:
        logger.info("Starting case details export process...")
        config = get_config()
        db = get_db_connection(config)
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