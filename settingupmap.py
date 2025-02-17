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

# Load data from the .mdb files
dfbusroute = load_mdb_data(mdb_files[0])
df2stopname = load_mdb_data(mdb_files[1])
dfdistance = load_mdb_data(mdb_files[2])

# Filter df2stopname for specific stop names
stop_namec_list = [
    "海富中心",
    "金鐘站",
    "力寶中心",
    "高等法院",
    "金鐘 - 太古廣場",
    "太古廣場",
    "金鐘（東）總站",
    "金鐘站 - 樂禮街",
    "金鐘 - 添馬街",
"金鐘 - 港鐵 D 出口",
"金鐘 - 港鐡 C 出口",
"金鐘 - 金鐘廊, 金鐘道",

"中環街市"
]

filtered_df = df2stopname[df2stopname["STOP_NAMEC"].isin(stop_namec_list)]

# Remove duplicate STOP_IDs
result = filtered_df[["STOP_ID", "STOP_NAMEE", "STOP_SEQ", "ROUTE_ID"]].drop_duplicates(subset="STOP_ID")
print(result)
# Add ON_STOP_ID to dfbusroute
dfbusroute = pd.merge(
    dfbusroute,
    df2stopname[["ROUTE_ID", "STOP_SEQ", "STOP_ID","ROUTE_SEQ"]].rename(columns={"STOP_ID": "ON_STOP_ID"}),
    left_on=["ROUTE_ID", "ON_SEQ","ROUTE_SEQ"],
    right_on=["ROUTE_ID", "STOP_SEQ","ROUTE_SEQ"],
    how="left"
)
print("First:")
print(dfbusroute)

# Add OFF_STOP_ID to dfbusroute
dfbusroute = pd.merge(
    dfbusroute,
    df2stopname[["ROUTE_ID", "STOP_SEQ", "STOP_ID","ROUTE_SEQ"]].rename(columns={"STOP_ID": "OFF_STOP_ID"}),
    left_on=["ROUTE_ID", "OFF_SEQ","ROUTE_SEQ"],
    right_on=["ROUTE_ID", "STOP_SEQ","ROUTE_SEQ"],
    how="left"
)
print("First:")
print(dfbusroute)
# Drop unnecessary columns
dfbusroute.drop(columns=["STOP_SEQ_x", "STOP_SEQ_y"], inplace=True)

# Add coordinates for ON_STOP_ID
dfbusroute = pd.merge(
    dfbusroute,
    dfdistance.rename(columns={"STOP_ID": "ON_STOP_ID", "X": "X_ON_STOPID", "Y": "Y_ON_STOPID"}),
    on="ON_STOP_ID",
    how="left"
)
print("First:")
print(dfbusroute)
# Add coordinates for OFF_STOP_ID
dfbusroute = pd.merge(
    dfbusroute,
    dfdistance.rename(columns={"STOP_ID": "OFF_STOP_ID", "X": "X_OFF_STOPID", "Y": "Y_OFF_STOPID"}),
    on="OFF_STOP_ID",
    how="left"
)

# Verify the row count to ensure no extra rows were added
assert len(dfbusroute) == len(dfbusroute), "Row count mismatch after merging!"

final_with_coordinates = dfbusroute 
# Add a new column for the Euclidean distance
final_with_coordinates["DISTANCE"] = np.sqrt(
    (final_with_coordinates["X_OFF_STOPID"] - final_with_coordinates["X_ON_STOPID"]) ** 2 +
    (final_with_coordinates["Y_OFF_STOPID"] - final_with_coordinates["Y_ON_STOPID"]) ** 2
)

# Display the updated DataFrame with the new DISTANCE column
print(final_with_coordinates)

final_with_coordinates.to_excel("C:\\Users\\bosco\\OneDrive\\桌面\\comp3522proj\\weare.xlsx", index=False)
print("finished")
# Export the DataFrame to an Excel file
destination_stops = set(filtered_df["STOP_ID"])

# Step 2: Extract Unique Stops
unique_stops = set(final_with_coordinates["ON_STOP_ID"]).union(set(final_with_coordinates["OFF_STOP_ID"]))



G = nx.DiGraph()  # Directed graph
for _, row in final_with_coordinates.iterrows():
    G.add_edge(
        row["ON_STOP_ID"],
        row["OFF_STOP_ID"],
        distance=row["DISTANCE"],
        price=row["PRICE"],
    )

# Step 4: Find Shortest Paths
# Step 4: Find Shortest Paths
result_data = []

for stop in unique_stops:
    if stop in destination_stops:
        continue  # Skip destination stops as starting points

    # Use Dijkstra's algorithm to find the shortest path to any destination stop
    try:
        # Compute shortest paths to all destinations
        shortest_path = None
        shortest_price = float("inf")
        shortest_distance = 0

        for dest in destination_stops:
            if nx.has_path(G, stop, dest):
                path_price = nx.shortest_path_length(G, source=stop, target=dest, weight="price")
                
                # Skip paths where the price exceeds $60
                if path_price > 60:
                    continue
                
                if path_price < shortest_price:
                    shortest_price = path_price
                    shortest_path = nx.shortest_path(G, source=stop, target=dest, weight="price")
                    # Calculate total distance for this path
                    shortest_distance = sum(
                        G[u][v]["distance"] for u, v in zip(shortest_path[:-1], shortest_path[1:])
                    )

        if shortest_path:
            # Add to results
            result_data.append({
                "ON_STOP_ID": stop,
                "OFF_STOP_ID": shortest_path[-1],
                "TOTAL_PRICE": shortest_price,
                "TOTAL_DISTANCE": shortest_distance,
            })

    except nx.NetworkXNoPath:
        # No valid path to any destination stop
        continue

# Step 5: Create Result DataFrame
result_df = pd.DataFrame(result_data)

# Display the result
print(result_df)
result_df.to_excel("C:\\Users\\bosco\\OneDrive\\桌面\\comp3522proj\\ad.xlsx", index=False)

print("Excel file has been saved!")