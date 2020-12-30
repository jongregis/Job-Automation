import psycopg2
from psycopg2 import OperationalError


def create_connection(db_name, db_user, db_password, db_host, db_port):
    connection = None
    try:
        connection = psycopg2.connect(
            database=db_name,
            user=db_user,
            password=db_password,
            host=db_host,
            port=db_port,
        )
        print("Connection to PostgreSQL DB successful")
    except OperationalError as e:
        print(f"The error '{e}' occurred")
    return connection


connection = create_connection("eca", '', '', "127.0.0.1", '5432')


def create_database(connection, query):
    connection.autocommit = True
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        print("Query executed successfully")
    except OperationalError as e:
        print(f"The error '{e}' occurred")

# create_database_query = "CREATE DATABASE eca"
# create_database(connection, create_database_query)


def execute_query(connection, query):
    connection.autocommit = True
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        print("Query executed successfully")
    except OperationalError as e:
        print(f"The error '{e}' occurred")


add_table = """
CREATE TABLE IF NOT EXISTS "E-Learning" (
    id SERIAL PRIMARY KEY,
    first TEXT NOT NULL,
    last TEXT NOT NULL,
    school TEXT NOT NULL,
    course TEXT NOT NULL,
    email TEXT,
    address TEXT,
    rep TEXT,
    invoice_number TEXT,
    start_date TEXT,
    amount TEXT
)
"""
# execute_query(connection, add_table)
