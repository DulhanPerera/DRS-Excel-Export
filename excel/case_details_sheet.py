from tables.case_details_table import create_case_details_table
# from tables.contact_details_table import create_contact_details_table
# from tables.remarks_table import create_remarks_table
# from tables.settlement_table import create_settlement_table
# from tables.settlement_plan_table import create_settlement_plan_table
# from tables.approve_table import create_approve_table
# from tables.case_status_table import create_case_status_table
# from tables.case_negotiation_table import create_case_negotiation_table
# from tables.ro_negotiation_table import create_ro_negotiation_table
# from tables.drc_table import create_drc_table
# from tables.recovery_officers_table import create_recovery_officers_table
from manipulation.data_fetcher import get_settlement_data, get_settlement_plan_data
import logging
import sys

logger = logging.getLogger('excel_data_writer')

def create_case_details_sheet(wb, case_data, db):
    """
    Create the Case Details sheet with all tables.
    """
    try:
        logger.info("Creating Case Details sheet...")
        ws = wb.active
        ws.title = "Case Details"
        
        # Define starting row and column for the first table
        start_row, start_col = 2, 1
        
        # Create the Case Details table
        next_row = create_case_details_table(ws, case_data, start_row, start_col, db)
        
        # Add a two-row gap between the tables
        gap_row = next_row + 2
        
        # # Create the Contact Details table
        # next_row = create_contact_details_table(ws, case_data, gap_row, start_col)
        
        # # Add a two-row gap between the Contact Info and Remarks tables
        # gap_row = next_row + 1
        
        # # Create the Remarks table
        # next_row = create_remarks_table(ws, case_data, gap_row, start_col)
        
        # # Add a two-row gap between the Remarks and Settlement tables
        # gap_row = next_row + 1
        
        # # Retrieve settlement data for the case
        # case_id = case_data.get("case_id")
        # settlements = get_settlement_data(db, case_id)
        
        # # Create the Settlement table if settlement data exists
        # if settlements:
        #     next_row = create_settlement_table(ws, settlements, gap_row, start_col)
        #     gap_row = next_row + 2  # Add a gap after the Settlement table
        
        # # Retrieve settlement plan data for the case
        # settlement_plans = get_settlement_plan_data(db, case_id)
        
        # # Create the Settlement Plan table if settlement plan data exists
        # if settlement_plans:
        #     next_row = create_settlement_plan_table(ws, settlement_plans, gap_row, start_col)
        #     gap_row = next_row + 2

        # # Retrieve approve data from the case_data
        # approve_data = case_data.get("approve", [])
        
        # # Create the Approve table if approve data exists
        # if approve_data:
        #     next_row = create_approve_table(ws, approve_data, gap_row, start_col) 
        #     gap_row = next_row + 2   
        
        # # Retrieve case status data from the case_data
        # case_status_data = case_data.get("case_status", [])
        
        # # Create the Case Status table if case status data exists
        # if case_status_data:
        #     next_row = create_case_status_table(ws, case_status_data, gap_row, start_col)
        #     gap_row = next_row + 2 

        # # Retrieve RO Negotiation data from the case_data
        # ro_negotiation_data = case_data.get("ro_negotiation", [])
        
        # # Create the RO Negotiation table if RO Negotiation data exists
        # if ro_negotiation_data:
        #     next_row = create_ro_negotiation_table(ws, ro_negotiation_data, gap_row, start_col)
        #     gap_row = next_row + 2

        # # Retrieve DRC data from the case_data
        # drc_data = case_data.get("drc", [])
        
        # # Create the DRC table if DRC data exists
        # if drc_data:
        #     next_row = create_drc_table(ws, drc_data, gap_row, start_col)    
        #     gap_row = next_row + 2

        #     # Retrieve Recovery Officers data from each DRC entry
        #     recovery_officers_data = []
        #     for drc in drc_data:
        #         recovery_officers_data.extend(drc.get("recovery_officers", []))
            
        #     # Create the Recovery Officers table if Recovery Officers data exists
        #     if recovery_officers_data:
        #         create_recovery_officers_table(ws, recovery_officers_data, gap_row, start_col)

        logger.info("Case Details sheet created successfully.")
        return ws
    except Exception as e:
        logger.error(f"Failed to create Case Details sheet: {e}")
        sys.exit(1)