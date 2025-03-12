<<<<<<<< HEAD:test/dataManipulation/data_fetcher.py
========
# data_retrieval.py
>>>>>>>> bc627a22ca3a26d51db54fbd15be6fd75d7fb12d:test/dataManipulation/data_retrieval.py
from bson import ObjectId
import logging

def get_arrears_band_value(db, current_arrears_band):
    """
    Retrieve the value for the given arrears band from the arrears_bands collection.
    """
    try:
        arrears_bands_collection = db["Arrears_bands"]
        arrears_bands_doc = arrears_bands_collection.find_one({})
        if arrears_bands_doc:
            return arrears_bands_doc.get(current_arrears_band)
        else:
            logging.warning("No arrears bands document found in the collection.")
            return None
    except Exception as e:
        logging.error(f"Failed to retrieve arrears band value: {e}")
        return None

def get_settlement_data(db, case_id):
    """
    Retrieve settlement data for the given case_id from the case_settlements collection.
    """
    try:
        settlements_collection = db["Case_settlements"]
        settlements = list(settlements_collection.find({"case_id": case_id}))
        if settlements:
            logging.info(f"Found {len(settlements)} settlement records for case_id: {case_id}")
        else:
            logging.warning(f"No settlement records found for case_id: {case_id}")
        return settlements
    except Exception as e:
        logging.error(f"Failed to retrieve settlement data: {e}")
        return []

def get_settlement_plan_data(db, case_id):
    """
    Retrieve settlement plan data for the given case_id from the case_settlements collection.
    """
    try:
        settlements_collection = db["Case_settlements"]
        settlements = list(settlements_collection.find({"case_id": case_id}))
        settlement_plans = []
        for settlement in settlements:
            if "settlement_plan" in settlement:
                for plan in settlement["settlement_plan"]:
                    # Add settlement_id to each plan
                    plan["settlement_id"] = settlement.get("settlement_id")
                    settlement_plans.append(plan)
        if settlement_plans:
            logging.info(f"Found {len(settlement_plans)} settlement plan records for case_id: {case_id}")
        else:
            logging.warning(f"No settlement plan records found for case_id: {case_id}")
        return settlement_plans
    except Exception as e:
        logging.error(f"Failed to retrieve settlement plan data: {e}")
        return []