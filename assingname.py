import pyodbc
import pandas as pd
import numpy as np
import networkx as nx
# Path to your .mdb files
mdb_files = [
    "C:\\Users\\bosco\\Downloads\\FARE_BUS.mdb",
    "C:\\Users\\bosco\\Downloads\\RSTOP_BUS.mdb",
    "C:\\Users\\bosco\\Downloads\\STOP_BUS.mdb",
]

# Function to load data from each .mdb file
def load_mdb_data(mdb_file_path):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={mdb_file_path};"
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    tables = [table.table_name for table in cursor.tables(tableType="TABLE")]
    print(f"Tables in {mdb_file_path}:", tables)
    table_name = tables[0]  # Assuming the table of interest is the first one
    query = f"SELECT * FROM [{table_name}]"
    df = pd.read_sql(query, conn)
    conn.close()
    return df

# Step 1: Load the Excel file into a DataFrame
excel_file_path = "C:\\Users\\bosco\\OneDrive\\桌面\\comp3522proj\\ad.xlsx"
excel_df = pd.read_excel(excel_file_path)
print(excel_df)
# Step 2: Load data from the .mdb file for df2stopname
df2stopname = load_mdb_data(mdb_files[1])
df2stopname = df2stopname.drop_duplicates(subset=["STOP_ID"])
# Step 3: Match STOP_NAMEC for ON_STOP_ID
# Merge to add ON_STOP_NAMEC
excel_df = excel_df.merge(
    df2stopname[["STOP_ID", "STOP_NAMEC","STOP_NAMEE"]].rename(columns={"STOP_ID": "ON_STOP_ID", "STOP_NAMEC": "ON_STOP_NAMEC", "STOP_NAMEE": "ON_STOP_NAMEE"}),
    on="ON_STOP_ID",
    how="left"
)
print(excel_df)
# Step 4: Match STOP_NAMEC for OFF_STOP_ID
# Merge to add OFF_STOP_NAMEC
excel_df = excel_df.merge(
    df2stopname[["STOP_ID", "STOP_NAMEC","STOP_NAMEE"]].rename(columns={"STOP_ID": "OFF_STOP_ID", "STOP_NAMEC": "OFF_STOP_NAMEC","STOP_NAMEE": "OFF_STOP_NAMEE"}),
    on="OFF_STOP_ID",
    how="left"
)

# Step 5: Display or save the updated DataFrame
print(excel_df)

# Save the updated DataFrame back to an Excel file
output_file_path = "C:\\Users\\bosco\\OneDrive\\桌面\\comp3522proj\\ad.xlsx"
excel_df.to_excel(output_file_path, index=False)

print("Updated Excel file has been saved!")