import numpy as np
import pyodbc
import xlwings as xw
import Utilities
import pandas as pd
import json

# Classes for Items and Sub Assemblies-----------------------------------------------------



class Universal_Functions():
    '''
    Class containing functions that can be utilised across the whole program
    '''
    @staticmethod     
    def duplicate_collapser(self, list_with_duplicate):
        '''
        Takes a list containing duplicate values, creates a count of how many
        times each value appears, and then returns a new list with each item
        once, with a multiplier of how many times it was present in the 
        original

        Usage: duplicate_collapser(self, list_with duplicates)

        Expected input:
            Class identifier 'self' and a list containing duplicate values.
        Expected output:
            A list containg values followed by multipliers and no duplicates.
        
        '''
        list_with_duplicate = [str(item) for item in list_with_duplicate]
        # Creating a set so there is a record of each unique value
        duplicates = set(list_with_duplicate)
        # Creating an empty count list to be used later
        count_list = []

        # Looping over each unique value
        for val in duplicates:
            count = 0

            # Adding each unique value to the list
            count_list.append(val)

            # Looping over each element in the original list and adding to the count where duplicates are found
            for element in list_with_duplicate:
                if element == val:
                    count += 1
            
            # Adding the number of appearances to the count list
            count_list.append(count)
        
        final_list = []

        # Looping over the count_list in increments of 2, so that every 1st and 2nd number is concatonated
        for i in range(0, len(count_list), 2):
            # Storing the values in final_list with a * to represent multiplication
            final_list.append(f"{count_list[i]} * {count_list[i + 1]}")
        return final_list
    
class Base_Sub_Assemblies():
    '''
    Class for the simplest manufactured component parts, containing attributes to be passed into the next database
    '''
    def __init__(self, description, raw_materials, size, build_time):
        # Description
        self.description = description

        # raw_materials is a tuple containing material type, cost. and density, all read off of the raw materials SQL table
        self.raw_materials = raw_materials[1]
        self.material_cost = raw_materials[2]

        self.size = size
        self.weight = size * raw_materials[3]
        
        
        self.build_time = build_time
        self.material_cost = size * self.material_cost

        # Overall cost is a calculated class variable
        self.overall_cost = self.material_cost * self.size + self.build_time * hourly_rate

    @property
    def __dict__(self):
        '''
        Setting the dictionary method to a dictionary of selected attributes, so that relevant ones
        are displayed

        Usage: object.__dict__

        Returns: Attributes Dictionary
        '''
        self.attributes = {
            "Description" : self.description,
            "Raw Materials" : self.raw_materials,
            "Material Cost" : self.material_cost,

            "Overall Cost" : self.overall_cost,
            "Weight" : self.weight,
            "Build Time" : self.build_time

        }
        return self.attributes
    
    def __repr__(self):
        return self.description
    
class Item(Base_Sub_Assemblies):
    '''
    Class for more complex sub assemblies/items, made from base sub assemblies
    '''
    def __init__(self, description, *sub_assemblies):
        self.description = description
        self.sub_assemblies = sub_assemblies
        self.cost = 0

        # Loops over the cost of each sub assembly element and then sums them
        for i in self.sub_assemblies:
            self.cost = self.cost + i.overall_cost



        # Uses duplicate_collapser to create a concise, readable list of the sub assemblies needed to make the item
        self.parts = Universal_Functions.duplicate_collapser(self, self.sub_assemblies)
    
    @property
    def __dict__(self):
        '''
        Setting the dictionary method to a dictionary of selected attributes, so that relevant ones
        are displayed

        Usage: object.__dict__

        Returns: Attributes Dictionary
        '''
        self.attributes = {
            "Description" : self.description,
            "Sub Assemblies" : self.parts,
            "Cost" : self.cost
        }
        return self.attributes
    def __repr__(self):
        return self.description
        
class Raw_Materials():
    '''
    Class for raw materials, needs to be developed further, but can utilise SQL Database inputs
    to store corresponding values within a single object
    '''
    def __init__(self, ProductName, cost_per_length):
        self.ProductName = ProductName
        self.cost_per_length = cost_per_length
    def __repr__(self):
        return self.description
    
# Set global variable so that it can be used by all classes in their methods
global hourly_rate

hourly_rate = 10



# Reads in an Excel sheet containing raw materials and some corresponding values, and stores them in a pandas data frame
df = pd.read_excel("C:\\Users\\AWilson\\Projects\\Python\\Database\\Raw_materials_.xlsx","sheet1")

# Connecting to the LocalDB, through 'pipe' (can be found using cmd), and selecting the ItemsDB, and assigning the connection to a variable
conn = pyodbc.connect(r'DRIVER={SQL Server};SERVER=np:\\.\pipe\LOCALDB#9E8F192B\tsql\query;DATABASE=ItemsDB;Trusted_Connection=yes;')
                                                            
# Creating a new cursor for writing/reading on the connection
cursor = conn.cursor()

# Writing SQL language commands across a multi-line string to preserve indentation, etc. ------

# Command inserts a product into the Raw Materials table, if it doesn't already exist, and takes given values that correspond to column headers in the table


sql = """
IF NOT EXISTS (SELECT 1 FROM RAW_MATERIALS WHERE ProductName =?) 
BEGIN
    INSERT INTO RAW_MATERIALS (ProductName, Price, LinearDensity) VALUES(?,?,?)
END"""

# Loops over the data frame read in from Excel, and then assigns them to the values variable which is then used with cursor.execute() to execute the query
# Each ? mark is a placeholder that must be filled, which is why there is a repeated df["ProductName"].iloc[i]

# Loops over the entire dataframe, and adds the elements to the SQL sheet
for i in range(len(df)):
    values = (df["ProductName"].iloc[i], df["ProductName"].iloc[i], float(df["Price/Length"].iloc[i]), float(df["linear_density"].iloc[i]))
    cursor.execute(sql, values) # Executes the command with the assigned variables
conn.commit() # Commits all pending commands to the database

#cursor.execute("SELECT ProductID, ProductName, Price, LinearDensity FROM RAW_MATERIALS WHERE ProductName = ?","Brass")
cursor.execute("SELECT ProductID, ProductName, Price, LinearDensity FROM RAW_MATERIALS") # Gets the column values from the SQL table
rows = cursor.fetchall() # Fetches them and assigns them to a variable

raw_material_list = [] 
for row in rows:
    #print(row)
    raw_material_list.append(tuple(row))

leg3 = Base_Sub_Assemblies("Leg", raw_materials = raw_material_list[0], size = 12, build_time = 3)
worktop = Base_Sub_Assemblies("Work top", raw_materials = raw_material_list[0], size = 12, build_time = 3)
table = Item("Table", leg3, leg3, leg3, leg3, worktop)
table_attributes = table.__dict__

table.cost = 200

print("DICTIONARY ",table.__dict__)

table_dict = table.__dict__


sql = """
IF NOT EXISTS (SELECT 1 FROM SUB_ASSEMBLIES_COMPLEX WHERE SubAssemblyName =?)
BEGIN
    INSERT INTO SUB_ASSEMBLIES_COMPLEX (SubAssemblyName, PartsRequired, ApproxPrice) VALUES(?,?,?)
END
"""
cursor.execute(sql, str(table_dict["Description"]), str(table_dict["Description"]), str(table_dict["Sub Assemblies"]), str(table_dict["Cost"]))
conn.commit()




print("\n\nSUB ASSEMBLIES ", table_dict["Sub Assemblies"])


df = pd.read_sql("SELECT * FROM SUB_ASSEMBLIES_COMPLEX", conn)
for row in rows:
    print(row)


for i in row:
    print(i)




conn.close()

print(leg3.__dict__)
print(df)
print(df.at[0,'SubAssemblyName'])

sub_assembly = df.at[0,'SubAssemblyName']
parts = df.at[0, 'PartsRequired']
approx_price = df.at[0, 'ApproxPrice']


print(table.__dict__)
parts = eval(parts)
for k in parts:
    print(k)
r'''
with xw.App(visible = False) as app:
    Utilities.safe_book_opener(app, "C:\Users\AWilson\Projects\Python\Items DataBase.xlsm", 'Item Database')
'''
    

items_data = pd.read_excel("C:\\Users\\AWilson\\Projects\\Python\\Items DataBase.xlsm")
print(items_data)