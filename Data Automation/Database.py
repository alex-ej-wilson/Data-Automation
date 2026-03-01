import numpy as np
import pyodbc
import xlwings as xw
import Utilities
import pandas as pd

# Classes for Items and Sub Assemblies-----------------------------------------------------
class Base_Sub_Assemblies():
    def __init__(self, description, raw_materials, size, build_time):
        self.description = description
        self.raw_materials = raw_materials[1]
        self.material_cost = raw_materials[2]
        self.size = size
        self.build_time = build_time
        self.weight = size * 3#density_list[raw_materials]
        self.material_cost = size * self.material_cost
        self.overall_cost = self.material_cost*self.size + self.build_time*hourly_rate
    def output(self):
        print(f"Description: {self.description}\nRaw Materials: {self.raw_materials}\nSize: {self.size} \nBuild Time: {self.build_time} \nWeight: {self.weight}")
    def cost(self):
        self.overall_cost = self.material_cost*self.size + self.build_time*hourly_rate
        return self.overall_cost
    def __repr__(self):
        return self.description
    
class Item(Base_Sub_Assemblies):
    def __init__(self, description, *sub_assemblies):
        self.description = description
        self.sub_assemblies = sub_assemblies
        self.cost = 0
        for i in self.sub_assemblies:
            self.cost = self.cost + i.overall_cost
        
    def output(self):
        print(f"{self.description} cost: £{self.cost}")

    def __repr__(self):
        return self.description
class Raw_Materials():
    def __init__(self, ProductName, cost_per_length):
        self.ProductName = ProductName
        self.cost_per_length = cost_per_length
    

global material_list, cost_per_size, hourly_rate

# DICTIONARIES CAN BE MADE INTO DATABASE TABLES
# Basic Values for testing classes ----------------------------------------------------------
density_list = {"Stainless Steel" : 1000, 
                 "Copper" : 700, 
                 "Brass" : 500, 
                 "Aluminium" : 300,
                }
cost_per_size = {"Stainless Steel" : 10, 
                 "Copper" : 7, 
                 "Brass" : 5, 
                 "Aluminium" : 3,
                }
hourly_rate = 10

# Execution and testing of basic classes ----------------------------------------------------
'''
leg = Base_Sub_Assemblies("Leg", "Stainless Steel", 2, 1)
leg.output()
leg.cost()
leg2 = Base_Sub_Assemblies("Leg", "Brass", 2.5, 1)

work_top = Base_Sub_Assemblies("Work Top", "Copper", 2, 0.5)
work_top2 = Base_Sub_Assemblies("Work Top", "Stainless Steel", 3, 0.6)
bench = Item("Bench", leg, leg, leg, leg, work_top)
#bench.output()


bench2 = Item("Bench", leg2, leg2, leg2, leg2, work_top2)
#print(vars(bench2))

print(bench.__dict__)
#print(work_top.__dict__)

double_bench = Item("Double Bench", bench, bench2)

print(double_bench.__dict__)

print(leg.__dict__)

for i in leg.__dict__:
    print(i)
'''
# Reading in from Excel --------------------------------------------------------------------------

df = pd.read_excel("C:\\Users\\AWilson\\Projects\\Python\\Database\\Raw_materials.xlsx","sheet1")
#print("DATA FRAME ",df)

last_row = df.iloc[-1]
#print(last_row["ProductName"])




# Using pyodbc to run SQL Queries ---------------------------------------------------------------------------------------------------
# Connect to LocalDB (Make sure LocalDB is installed!)

conn = pyodbc.connect(r'DRIVER={SQL Server};SERVER=np:\\.\pipe\LOCALDB#802C0F5E\tsql\query;DATABASE=ItemsDB;Trusted_Connection=yes;')
cursor = conn.cursor()

# Update a record
sql = """
IF NOT EXISTS (SELECT 1 FROM RAW_MATERIALS WHERE ProductName =?) 
BEGIN
    INSERT INTO RAW_MATERIALS (ProductName, Price) VALUES(?,?)
END"""
# Loops over the data frame read in from Excel, and then assigns them to the values variable which is then used with cursor.execute() to execute the query
for i in range(len(df)):
    #print(f'{i}: {df["ProductName"].iloc[i]}, {df["Price/Length"].iloc[i]}')
    values = (df["ProductName"].iloc[i], df["ProductName"].iloc[i], float(df["Price/Length"].iloc[i]))
    cursor.execute(sql, values)
conn.commit()
cursor.execute("SELECT ProductID, ProductName, Price FROM RAW_MATERIALS WHERE ProductName = ?","Brass")
#cursor.execute("SELECT ProductID, ProductName, Price FROM RAW_MATERIALS")
rows = cursor.fetchall()

# Print the rows

for row in rows:
    print(row)
    print(f"{row.ProductName}, £{row.Price}, {row.LinearDensity}")

# Close the connection
# Commit changes and close

conn.close()

print("LocalDB updated successfully!")
raw_material = tuple(row)
#print("ROW", raw_material[2])
leg3 = Base_Sub_Assemblies("Leg", raw_materials = raw_material, size = 12, build_time = 3)
print(leg3.__dict__)
table = Item("Table", leg3, leg3, leg3, leg3)
print(table.__dict__)