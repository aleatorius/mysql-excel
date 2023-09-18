from pymongo import MongoClient
def get_database(database):
    #CONNECTION_STRING = "mongodb://localhost:27017"
    CONNECTION_STRING = "mongodb://129.241.18.103:27017/"
    client = MongoClient(CONNECTION_STRING)
    return client[database]


if __name__ == "__main__":   
    database = 'calst'
    # Get the database
    dbname = get_database()

