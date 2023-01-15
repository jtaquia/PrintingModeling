# -*- coding: utf-8 -*-
"""factory_planning_2.ipynb


---
## Python Implementation

We import the Gurobi Python Module and other Python libraries.
"""

import numpy as np
import pandas as pd

import gurobipy as gp
from gurobipy import GRB

# tested with Python 3.7.0 & Gurobi 9.0

"""## Input Data
We define all the input data of the model.
"""

# Parameters

products = ["Prod1", "Prod2", "Prod3", "Prod4", "Prod5", "Prod6", "Prod7","Prod8","Prod9","Prod10"]
machines = ["CTP", "HEIDELBERG_OFFSET", "HEIDELBERG_GTO", "Plastificadora","Guillotina"]
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

profit = {"Prod1":9, "Prod2":10, "Prod3":6, "Prod4":8, "Prod5":4, "Prod6":11, "Prod7":9,"Prod8":3,"Prod9":4,"Prod10":11}


time_req = {
    "CTP": {    "Prod1": 0.17, "Prod2": 0.17, "Prod3": 0.17,
                    "Prod4": 0.17 , "Prod5": 0.17, "Prod6":0 ,
                    "Prod7":0 ,"Prod8": 0.17, "Prod9":0 ,"Prod10":0 ,},
    "HEIDELBERG_OFFSET": {  "Prod1": 0, "Prod2": 0.13, "Prod3": 0,
                    "Prod4": 0.06 , "Prod5": 0, "Prod6":0 ,
                    "Prod7":0 ,"Prod8": 0, "Prod9":0 ,"Prod10":0 ,},
    "HEIDELBERG_GTO": { "Prod1": 0.10, "Prod2": 0.20, "Prod3": 0.05, 
                   "Prod4":0, "Prod5": 0.05 , "Prod6": 0.10, "Prod7": 0.10,
                    "Prod8": 0.05, "Prod9": 0.10, "Prod10": 0.10,},
    "Plastificadora": {   "Prod1": 0, "Prod2": 0, "Prod3": 0,
                    "Prod4": 0 , "Prod5": 0, "Prod6":0 ,
                    "Prod7":0 ,"Prod8": 0.25, "Prod9":0 ,"Prod10":0 ,},
    "Guillotina": {   "Prod1": 0.03, "Prod2": 0.07, "Prod3": 0, "Prod4": 0.05 , 
                    "Prod5": 0.13 , "Prod6": 0.03, "Prod7": 0.03,
                    "Prod8": 0, "Prod9": 0.03, "Prod10": 0.03,}
}

# number of each machine available
installed = {"CTP":1, "HEIDELBERG_OFFSET":1, "HEIDELBERG_GTO":3, "Plastificadora":1, "Guillotina":1} 

# number of machines that need to be under maintenance
down_req = {"CTP":1, "HEIDELBERG_OFFSET":1, "HEIDELBERG_GTO":0, "Plastificadora":0, "Guillotina":1} 

################################################
################################################

longitud= len(products)
longitud
cantidad=len(months)

months[0]*longitud


##############################################
def tuplas():
    mes=[]
    mes_products=[]
    for i in range(cantidad):
        mes=[months[i]]*longitud
        for j in zip(mes, products):
            #print(j)    
            mes_products.append(j)
    
    return mes_products


tuplas()
#################################################
import os
#import pandas as pd
os.chdir("D://DOCUMENTOS//UNIVERSIDAD_DE_LIMA//PROYECTOS_2//2022 1//1022//imprentaMoyobamba//modeloImprenta")
os.getcwd()

tabla1= pd.read_excel( "ImprentaMoyobamba-Excel.xlsx",sheet_name="Hoja2",header=0,index_col=0)
tabla1
vector2=tabla1.values.tolist()
cantidad2=len(vector2)
cantidad3=len(vector2[0])

def valor():
    ValorMes_products=[]
    for i in range(cantidad2):
        #print(i)
        for j in range(cantidad3):
           # print(vector2[i][j])    
            ValorMes_products.append(vector2[i][j])
    return ValorMes_products


valor()
##########################################
secuenciaValores=[]
dicc2=[]
secuenciaValores=valor()
dicc2=tuplas()
dicc2[0]

def zipDemanda(key_list,value_list):    
    diccionarioB = dict(zip(key_list, value_list))
    return diccionarioB

#zipDemanda(dicc2,secuenciaValores)



################################################
################################################



# market limitation of sells
max_sales = zipDemanda(dicc2,secuenciaValores)
holding_cost = 0.5
max_inventory = 100
store_target = 50
hours_per_month = 2*8*24

"""## Model Deployment
We create a model and the variables. We set the UpdateMode parameter to 1 (which simplifies the code – see the documentation for more details). For each product (seven kinds of products) and each time period (month), we will create variables for the amount of which products will get manufactured, held, and sold. In each month, there is an upper limit on the amount of each product that can be sold. This is due to market limitations. For each type of machine and each month we create a variable d, which tells us how many machines are down in this month of this type.
"""

factory = gp.Model('Factory Planning II')

make = factory.addVars(months, products, name="Make") # quantity manufactured
store = factory.addVars(months, products, ub=max_inventory, name="Store") # quantity stored
sell = factory.addVars(months, products, ub=max_sales, name="Sell") # quantity sold
repair = factory.addVars(months, machines, vtype=GRB.INTEGER, ub=down_req, name="Repair") # number of machines down

"""Next, we insert the constraints.
The balance constraints ensure that the amount of product that is in the storage in the prior month and the amount that gets manufactured equals the amount that is sold and held for each product in the current month. This ensures that all products in the model are manufactured in some month. The initial storage is empty.
"""

#1. Initial Balance
Balance0 = factory.addConstrs((make[months[0], product] == sell[months[0], product] 
                  + store[months[0], product] for product in products), name="Initial_Balance")
    
#2. Balance
Balance = factory.addConstrs((store[months[months.index(month) -1], product] + 
                make[month, product] == sell[month, product] + store[month, product] 
                for product in products for month in months 
                if month != months[0]), name="Balance")

"""The endstore constraints force that at the end of the last month the storage contains the specified amount of each product."""

#3. Inventory Target
TargetInv = factory.addConstrs((store[months[-1], product] == store_target for product in products),  name="End_Balance")

"""The capacity constraints ensure that for each month the time all products require on a certain kind of machine is lower or equal than the available hours for that machine in that month multiplied by the number of available machines in that month. Each product requires some machine hours on different machines. Each machine is down in one or more months due to maintenance, so the number and types of available machines varies per month. There can be multiple machines per machine type."""

#4. Machine Capacity
MachineCap = factory.addConstrs((gp.quicksum(time_req[machine][product] * make[month, product]
                             for product in time_req[machine])
                    <= hours_per_month * (installed[machine] - repair[month, machine])
                    for machine in machines for month in months),
                   name = "Capacity")

"""The maintenance constraints ensure that the specified number and types of machines are down due maintenance in some month. Which month a machine is down is now part of the optimization."""

#5. Maintenance

Maintenance = factory.addConstrs((repair.sum('*', machine) == down_req[machine] for machine in machines), "Maintenance")

"""The objective is to maximize the profit of the company, which consists of the profit for each product minus cost for storing the unsold products. This can be stated as:"""

#0. Objective Function
obj = gp.quicksum(profit[product] * sell[month, product] -  holding_cost * store[month, product]  
               for month in months for product in products)

factory.setObjective(obj, GRB.MAXIMIZE)

"""Next, we start the optimization and Gurobi finds the optimal solution."""

factory.optimize()

"""---
## Analysis

The result of the optimization model shows that the maximum profit we can achieve is $\$108,855.00$. This is an increase of $\$15,139.82$ over the course of six months compared to the Factory Planning I example as a result of being able to pick the maintenance schedule as opposed to having a fixed one. Let's see the solution that achieves that optimal result. 

### Production Plan
This plan determines the amount of each product to make at each period of the planning horizon. For example, in February we make 600 units of product Prod1.
"""

rows = months.copy()
columns = products.copy()
make_plan = pd.DataFrame(columns=columns, index=rows, data=0.0)

for month, product in make.keys():
    if (abs(make[month, product].x) > 1e-6):
        make_plan.loc[month, product] = np.round(make[month, product].x, 1)
make_plan

"""### Sales Plan
This plan defines the amount of each product to sell at each period of the planning horizon. For example, in February we sell 600 units of product Prod1.
"""

rows = months.copy()
columns = products.copy()
sell_plan = pd.DataFrame(columns=columns, index=rows, data=0.0)

for month, product in sell.keys():
    if (abs(sell[month, product].x) > 1e-6):
        sell_plan.loc[month, product] = np.round(sell[month, product].x, 1)
sell_plan
dfasigna = pd.DataFrame(data=sell_plan) #, index=[0])
dfasigna = (dfasigna.T)
dfasigna.to_excel('sell_plan.xlsx', sheet_name='sheet1')#, index=False)
    
"""### Inventory Plan
This plan reflects the amount of product in inventory at the end of each period of the planning horizon. For example, at the end of February we have zero units of Prod1 in inventory.
"""

rows = months.copy()
columns = products.copy()
store_plan = pd.DataFrame(columns=columns, index=rows, data=0.0)

for month, product in store.keys():
    if (abs(store[month, product].x) > 1e-6):
        store_plan.loc[month, product] = np.round(store[month, product].x, 1)
store_plan
dfasigna = pd.DataFrame(data=store_plan)
dfasigna = (dfasigna.T)
dfasigna.to_excel('Store_Plan.xlsx', sheet_name='sheet1')#, index=False)
    


"""### Maintenance Plan
This plan shows the maintenance plan for each period of the planning horizon. For example, 2 machines of type grinder will be down for maintenance in April.
"""

rows = months.copy()
columns = machines.copy()
repair_plan = pd.DataFrame(columns=columns, index=rows, data=0.0)

for month, machine in repair.keys():
    if (abs(repair[month, machine].x) > 1e-6):
        repair_plan.loc[month, machine] = repair[month, machine].x
repair_plan





"""**Note:** If you want to write your solution to a file, rather than print it to the terminal, you can use the model.write() command. An example implementation is:

`factory.write("factory-planning-2-output.sol")`

---
## References

H. Paul Williams, Model Building in Mathematical Programming, fifth edition.

Copyright &copy; 2020 Gurobi Optimization, LLC
"""

