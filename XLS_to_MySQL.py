# Python script for inserting data from an XLS file to two MYSQL tables, one which has a foreign key constraint to an auto-increment to a parent table.
# - Based on a database model that uses a parent table with an auto-increment column and a child table which is foreign key constrained to the increment value on the parent table
# - deals with null values by replacing the \N character that is interpreted in MySQL as a null
# - places quotes around string values
# - generates a query based on the number of columns specified in the sheet_idx variable


import xlrd
import pymysql
import os

create_tables = """

SET FOREIGN_KEY_CHECKS = 0;

CREATE TABLE IF NOT EXISTS parent ( col1 FLOAT, col2 FLOAT, col3 INTEGER(8), col4 INTEGER(8), ID INTEGER(8) UNIQUE NOT NULL AUTO_INCREMENT,  PRIMARY KEY (ID) );

CREATE TABLE IF NOT EXISTS child ( col1 VARCHAR(45), col2 VARCHAR(45), col3 INTEGER(8), ID INTEGER(8) UNIQUE NOT NULL, PRIMARY KEY (ID), FOREIGN KEY (ID) REFERENCES parent(ID) );

SET FOREIGN_KEY_CHECKS = 1;

"""

row_remove = """

SET SQL_SAFE_UPDATES = 0;

DELETE FROM child
    WHERE col1 is NULL
        AND col2 is NULL
            AND col3 is NULL;

SET SQL_SAFE_UPDATES = 1;

"""

database = pymysql.connect(host="localhost", user="root",
                           passwd="yourpass", charset="utf8", db="yourdb")      	# Establishes the database connection

os.chdir(r"C:\Yourpath")                                                			# Choose working directory
path = os.getcwd()                                                                  # Store directory path
file = path + "\\" + "XLS_to_MYSQL.xls"                                          	# Path to file

book = xlrd.open_workbook(file)                                                     # Opens the workbook
sheet = book.sheet_by_name('main')                                                  # Selects the sheet based on main
sheet_idx = [[0, 4, 'parent'], [5, 7, 'child']]                                     # Sheet index and table names to be generated, can be modified to accommodate more tables

def stringGen(num_val, sheet_name):                                                 # Function to generate the string used for query

    num_str = ("INSERT INTO {} VALUES ({}")											# Start of insert statement
    for n in range(num_val):                                                        # Adds number of formats based on the column index specified in sheet_idx
        num_add = ", {}"
        num_str = num_str + num_add

    if sheet_name == 'parent':                                                      # End of query format for the master column
        query_str = num_str + ")"
    else:                                                                           # End of query format for the child column including auto-increment
        query_str = num_str + ", LAST_INSERT_ID())"

    return query_str                                                                # Returns completed query


def row_access(sheets, row):

    var_lst = []
    num_val = int(sheets[1]) - int(sheets[0])                       # Numerical value used to create the number of entries for query formatting
    for col in range(int(sheets[0]), (int(sheets[1]) + 1)):         # Iterates across the cells in the row
        val = sheet.cell(row, col).value                            # Pulls out the value from the cell in XLS

        if not val:                                                 # If the cell is empty, substitute with \N
            val = r"\N"
        if isinstance(val, str):                                    # If the entry is a string
            if val[0] != '\\':                                      # And not starting with '\'
                val = "'" + val + "'"                               # then place quotes around the field

        var_lst.append(val)
    var_lst = tuple(var_lst)

    return stringGen(num_val, sheets[2]), var_lst


with database.cursor() as cursor:
    cursor.execute(create_tables)                                   # Creates tables for updating
    for r in range(1, sheet.nrows):                                 # Iterates down rows, skipping headers

        for sx in sheet_idx:                                        # For each cell indexes outlined in variable
            query_str, var_lst = row_access(sx, r)                  # Calls function to retrieve data from XLS rows
            query = query_str.format(sx[2], *var_lst)
            cursor.execute(query_str.format(sx[2], *var_lst))       # Executes query for each row
            print(query)                                            # Console output

    cursor.execute(row_remove)                                      # Row clean up to delete empty rows
    database.commit()

database.close()