"""
DEV NOTES: ARBAZ QURESHI (TTC4624)
This is a test script, which fetches data from ServiceNow api, and simply uploads it to a table in DB
"""
import os
import json
import requests
import pandas as pd
import mysql
from mysql import connector
from dotenv import load_dotenv


# ══════════════════════════════════════════════════════════
#  FUNCTION 1: FETCH DATA FROM SERVICENOW API
# ══════════════════════════════════════════════════════════
def fetch_api_data() -> list:
    """
    Fetches incident data from ServiceNow API.
    Returns: list of incidents (dicts)
    """
    api_key  = os.getenv("api_token")
    api_host = os.getenv("api_host")
    url      = "https://kroger.service-now.com/api/now/table/incident"

    headers = {
        "Authorization"  : api_key,
        "Content-Type"   : "application/json",
        "Accept"         : "application/json"
    }

    querystring = {
        "sysparm_query": (
            "assignment_groupIN"
            "7cd3576c1b1a4d508611caa3604bcbfa,"   # APP-DIG-ORE2.0-Platform
            "dc8c9dfe1b160d10684965f3604bcbea"    # APP-DIG-ORE2.0-eCommInnovationAPI
            "^active=true"
            "^stateIN4,5"                          # In Queue OR Assigned
            "^ORDERBYDESCsys_created_on"
        ),
        "sysparm_fields"                : "number,cmdb_ci,description",
        "sysparm_limit"                 : "10",
        "sysparm_exclude_reference_link": "true"
    }

    response  = requests.get(url, headers=headers, params=querystring)
    incidents = response.json()['result']

    print("\n")
    print(incidents)
    print(f"\n📋 Total Incidents Fetched: {len(incidents)}")

    return incidents                                # ← sends incidents back to main


# ══════════════════════════════════════════════════════════
#  FUNCTION 2: CONNECT TO DATABASE
# ══════════════════════════════════════════════════════════
def connect_to_database() -> tuple:
    """
    Connects to MySQL database using env variables.
    Returns: tuple of (connection, cursor)
    """
    MYSQL_HOST     = os.getenv("MYSQL_HOST")
    MYSQL_PORT     = os.getenv("MYSQL_PORT")
    MYSQL_USER     = os.getenv("MYSQL_USER")
    MYSQL_PASSWORD = os.getenv("MYSQL_PASSWORD")
    MYSQL_DATABASE = os.getenv("MYSQL_DATABASE")

    server_connection = mysql.connector.connect(
        host              = MYSQL_HOST,
        port              = MYSQL_PORT,
        user              = MYSQL_USER,
        password          = MYSQL_PASSWORD,
        database          = MYSQL_DATABASE,
        autocommit        = False,
        connection_timeout= 20,
        raise_on_warnings = True
    )

    if server_connection.is_connected():
        print("\n✅ Connected to DATABASE!")
    else:
        print("\n❌ Connection Failed!")
        raise SystemExit("Could not connect to database!")

    cursor = server_connection.cursor()

    return server_connection, cursor               # ← sends both back to main


# ══════════════════════════════════════════════════════════
#  FUNCTION 3: VALIDATE TABLE EXISTS IN DATABASE
# ══════════════════════════════════════════════════════════
def validate_table(cursor, sql_table: str) -> bool:
    """
    Checks if the target table exists in the database.
    Returns: True if table exists, raises error if not
    """
    cursor.execute("SHOW TABLES LIKE %s", (sql_table,))

    if cursor.fetchone() is None:
        raise SystemExit(f"\n❌ The Table '{sql_table}' was not found!")
    else:
        print(f"\n✅ Using The Table '{sql_table}'")

    return True                                    # ← sends True back to main


# ══════════════════════════════════════════════════════════
#  FUNCTION 4: UPLOAD / INSERT DATA TO DATABASE
# ══════════════════════════════════════════════════════════
def upload_data_to_db(connection, cursor, incidents: list, sql_table: str) -> tuple:
    """
    Inserts incidents into the database table.
    Returns: tuple of (success_count, error_count, success_list, error_list)
    """
    print(f"\n📋 Rows to upsert: {len(incidents)}")

    success = 0
    error   = 0
    suc     = []    # successful incident numbers
    ex      = []    # exceptions/errors

    for inc in incidents:
        sql    = f"INSERT INTO {sql_table} (inc_num, ci, description) VALUES (%s, %s, %s)"
        values = (inc['number'], inc['cmdb_ci'], inc['description'])

        try:
            cursor.execute(sql, values)
            success += 1
            suc.append(inc['number'])
        except Exception as e:
            connection.rollback()
            ex.append(e)
            error += 1

    connection.commit()

    return success, error, suc, ex                 # ← sends results back to main


# ══════════════════════════════════════════════════════════
#  FUNCTION 5: PRINT UPLOAD RESULTS / SUMMARY
# ══════════════════════════════════════════════════════════
def print_upload_summary(success: int, error: int, suc: list, ex: list) -> None:
    """
    Prints the summary of upload results.
    """
    if success == 0 and error > 0:
        print(f"\n❌ [ERROR] Rolled back due to following {error} errors:")
        for e in ex:
            print(f"   → {e}")

    elif success > 0:
        print(f"\n✅ [SUCCESS] Successfully inserted following {success} incidents:")
        for inc_num in suc:
            print(f"   → {inc_num}")

        if error > 0:
            print(f"\n⚠️  [WARNING] {error} incidents failed to insert:")
            for e in ex:
                print(f"   → {e}")


# ══════════════════════════════════════════════════════════
#  FUNCTION 6: VERIFY DATA IN DATABASE
# ══════════════════════════════════════════════════════════
def verify_db_data(cursor, sql_table: str) -> None:
    """
    Fetches and prints all records from the table for verification.
    """
    cursor.execute(f"SELECT * FROM {sql_table}")
    result = cursor.fetchall()

    print(f"\n📋 Total Records in '{sql_table}': {len(result)}")
    print("-" * 60)

    for record in result:
        print(f"  INC Number : {record[0]}")

    print("-" * 60)


# ══════════════════════════════════════════════════════════
#  FUNCTION 7: CLOSE DATABASE CONNECTION
# ══════════════════════════════════════════════════════════
def close_db_connection(connection, cursor) -> None:
    """
    Closes cursor and database connection safely.
    """
    cursor.close()
    connection.close()
    print("\n✅ MySQL Connection Closed!")
    print("=" * 60)


# ══════════════════════════════════════════════════════════
#  MAIN - ORCHESTRATES ALL FUNCTIONS
# ══════════════════════════════════════════════════════════
def main():
    sql_table = "parent_child_usecase_try2"

    # Step 1: Fetch data from API
    incidents = fetch_api_data()

    # Step 2: Connect to database
    connection, cursor = connect_to_database()

    # Step 3: Validate table exists
    validate_table(cursor, sql_table)

    # Step 4: Upload data to DB
    success, error, suc, ex = upload_data_to_db(connection, cursor, incidents, sql_table)

    # Step 5: Print summary
    print_upload_summary(success, error, suc, ex)

    # Step 6: Verify data in DB
    verify_db_data(cursor, sql_table)

    # Step 7: Close connection
    close_db_connection(connection, cursor)


if __name__ == '__main__':
    load_dotenv()
    main()