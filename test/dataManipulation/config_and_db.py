<<<<<<<< HEAD:test/dataManipulation/config_loader.py
# config_loader.py
========
# config_and_db.py
>>>>>>>> bc627a22ca3a26d51db54fbd15be6fd75d7fb12d:test/dataManipulation/config_and_db.py
import configparser
import logging
import os
import sys
from pymongo import MongoClient
<<<<<<<< HEAD:test/dataManipulation/config_loader.py
import logging

def connect_db(config):
    """
    Connect to the MongoDB database using the configuration.
    """
    try:
        client = MongoClient(config['DATABASE']['MONGO_URI'])
        db = client[config['DATABASE']['DB_NAME']]
        logging.info("Successfully connected to the database.")
        return db
    except Exception as e:
        logging.error(f"Failed to connect to the database: {e}")
        sys.exit(1)

========
>>>>>>>> bc627a22ca3a26d51db54fbd15be6fd75d7fb12d:test/dataManipulation/config_and_db.py

def get_config():
    """
    Load and return the configuration from the Config.ini file.
    """
    try:
        config = configparser.ConfigParser()
        config.read(os.path.join(os.path.dirname(__file__), '../Config/Config.ini'))
        if not config.sections():
            logging.error("Configuration file is empty or not found.")
            sys.exit(1)
        return config
    except Exception as e:
        logging.error(f"Failed to load configuration: {e}")
<<<<<<<< HEAD:test/dataManipulation/config_loader.py
========
        sys.exit(1)

def connect_db(config):
    """
    Connect to the MongoDB database using the configuration.
    """
    try:
        client = MongoClient(config['DATABASE']['MONGO_URI'])
        db = client[config['DATABASE']['DB_NAME']]
        logging.info("Successfully connected to the database.")
        return db
    except Exception as e:
        logging.error(f"Failed to connect to the database: {e}")
>>>>>>>> bc627a22ca3a26d51db54fbd15be6fd75d7fb12d:test/dataManipulation/config_and_db.py
        sys.exit(1)