# config_loader.py
import configparser
import logging
import os
import sys
from pymongo import MongoClient
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
        sys.exit(1)