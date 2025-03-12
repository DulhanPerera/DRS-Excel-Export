<<<<<<<< HEAD:test/dataManipulation/create_sheet.py
from export_details import export_case_details
from config_loader import get_config, connect_db
import logging


def create_excel_sheet():
========
# main.py
import logging
import logging.config
import sys  # Add this import
from config_and_db import get_config, connect_db
from excel_export import export_case_details

def main():
>>>>>>>> bc627a22ca3a26d51db54fbd15be6fd75d7fb12d:test/dataManipulation/main.py
    """
    Main function to execute the case details export process.
    """
    try:
<<<<<<<< HEAD:test/dataManipulation/create_sheet.py
        logging.info("Starting case details export process...")
========
        logging.config.fileConfig('Config/logger/loggers.ini')
        logger = logging.getLogger('excel_data_writer')
        
        logger.info("Starting case details export process...")
>>>>>>>> bc627a22ca3a26d51db54fbd15be6fd75d7fb12d:test/dataManipulation/main.py
        config = get_config()
        db = connect_db(config)
        incident_id = 78910  # Replace with the actual incident ID you want to export
        export_path = config['EXCEL_EXPORT_PATHS']['WIN_DB']
        collection_name = config['COLLECTIONS']['CASE_DETAIL_COLLECTION_NAME']
        
        export_case_details(db, incident_id, export_path, collection_name)
        
<<<<<<<< HEAD:test/dataManipulation/create_sheet.py
        logging.info("Case details export process completed.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

# if __name__ == "__main__":
#     create_excel_sheet()
========
        logger.info("Case details export process completed.")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
>>>>>>>> bc627a22ca3a26d51db54fbd15be6fd75d7fb12d:test/dataManipulation/main.py
