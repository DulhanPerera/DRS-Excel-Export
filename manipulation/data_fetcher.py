import logging
from datetime import datetime, timedelta
from bson import ObjectId

logger = logging.getLogger('excel_data_writer')

def fetch_incidents(db, filters=None):
    """Fetch incidents from Incident_log collection with precise filtering"""
    try:
        collection = db["Incident_log"]
        query = {}
        
        if filters:
            if 'action' in filters:
                action = filters['action']
                if action == 'collect arrears and CPE':
                    query["Actions"] = {"$regex": r'^collect arrears and CPE$'}
                elif action in ['collect arrears', 'collect CPE']:
                    query["Actions"] = action
            
            if 'status' in filters:
                status = filters['status']
                valid_statuses = ['Incident Open', 'Incident Close', 'Incident Reject']
                
                if isinstance(status, list):
                    query["Incident_Status"] = {"$in": [s for s in status if s in valid_statuses]}
                elif status in valid_statuses:
                    query["Incident_Status"] = status
            
            if 'date_range' in filters:
                start_date, end_date = filters['date_range']
                query["Created_Dtm"] = {"$gte": start_date, "$lte": end_date}
        
        logger.info(f"Executing query: {query}")
        incidents = list(collection.find(query))
        logger.info(f"Found {len(incidents)} matching incidents")
        return incidents
    
    except Exception as e:
        logger.error(f"Error fetching incidents: {str(e)}")
        return []

def fetch_template_task(db, task_id):
    """Fetch template task from Template_Tasks collection"""
    try:
        collection = db["Templete_tasks"]
        task = collection.find_one({"Template_Task_Id": task_id})
        return task
    except Exception as e:
        logger.error(f"Error fetching template task {task_id}: {str(e)}")
        return None
    
