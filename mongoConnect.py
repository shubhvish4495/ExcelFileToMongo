from pymongo import MongoClient
import configuration as cfg
from pprint import pprint

def connect():
    try:
    # connect to MongoDB, change the << MONGODB URL >> to reflect your own connection string
        client = MongoClient(cfg.MONGO_HOST, cfg.MONGO_PORT)
        db=client[cfg.MONGO_DBNAME]
        # Issue the serverStatus command and print the results

        serverStatusResult=db.command("serverStatus")
        print("Connected to db successfully")
        return db
    except Exception as e:
        print("Error in connection : \n", e)