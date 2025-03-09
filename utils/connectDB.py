import configparser
from pymongo import MongoClient
import os

def get_db_connection():
    config_path = os.path.join(os.path.dirname(__file__), "../Config/database/DB_Config.ini")

    # Check if configuration file exists
    if not os.path.exists(config_path):
        print(f"Configuration file '{config_path}' not found.")
        return None

    # Read the configuration file
    config = configparser.ConfigParser()
    config.read(config_path)

    # Ensure 'DATABASE' section exists
    if 'DATABASE' not in config:
        print("'DATABASE' section not found in DB_Config.ini")
        return None

    # Retrieve values
    mongo_uri = config['DATABASE'].get('MONGO_URI', '').strip()
    db_name = config['DATABASE'].get('DB_NAME', '').strip()

    if not mongo_uri or not db_name:
        print("Missing MONGO_URI or DB_NAME in DB_Config.ini")
        return None

    try:
        # Connect to MongoDB
        client = MongoClient(mongo_uri)
        db = client[db_name]
        print(f"Connected to MongoDB successfully | dataBase name:{db_name}")
        return db
    except Exception as e:
        print(f"Error connecting to MongoDB: {e}")
        return None


if __name__ == "__main__":
    db_connection = get_db_connection()

    # Fix: Explicitly check if the connection is None
    if db_connection is not None:
        print("Database connection established.")
    else:
        print("Database connection failed.")
