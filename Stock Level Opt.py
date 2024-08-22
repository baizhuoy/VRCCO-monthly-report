import numpy as np, pandas as pd
from pandas import DataFrame
import re
import math
import openpyxl

import gurobipy as gp
from gurobipy import *
from gurobipy import GRB

# input data from AI demand forecasting model and product info
demand_path=input("What is the demand file path: (like \\VRCCO-24\\Users\Michael\Desktop\INV EXCEL FILES\Demand Forecast.xlsx) ").strip() or r"\\VRCCO-24\\Users\Michael\Desktop\INV EXCEL FILES\Demand Forecast.xlsx"
product_path=input("What is the Master List File path: (like \\VRCCO-24\\Users\Michael\Desktop\INV EXCEL FILES\Monthly Inv Report.xlsx) ").strip() or r"\\VRCCO-24\\Users\Michael\Desktop\INV EXCEL FILES\Monthly Inv Report.xlsx"
stock_level=input('what is the Stock Level File path').strip()or r"\\VRCCO-24\\Users\Michael\Desktop\INV EXCEL FILES\Stock level.xlsx"

demand=pd.read_excel(demand_path,sheet_name='Forecast',converters={'SKU':str})
inv=pd.read_excel(demand_path,sheet_name='History',converters={'SKU':str})
producttt=pd.read_excel(product_path,sheet_name=None)
last_sheet = list(producttt.keys())[-1]
product=pd.read_excel(product_path,sheet_name=last_sheet,converters={'SKU':str})

# create model
m = gp.Model("Inventory_Optimization")

# add parameters
Si = {}  # start inventory
Dit = {}  # daily demand
P = []  # product set
SSpi = {}  # max inventory
Ri = {}  # safety stock
Qi = {}  # unit price
Ki = {}  # unit quantity
T = 7
ordering_cost = 200
lead_time = 2  # purchase lead time
M = 1000000000

# input data
for a, b in enumerate(inv['Product']):
    Si[b] = int(inv.iloc[a, -1])
for t in range(T):
    for a, b in enumerate(demand['Product']):
        Dit[b,t] = math.ceil(float(demand.iloc[a, -1] / 7))
for a in enumerate(product['Product']):
    prod=str(a[1])
    P.append(prod)
    SSpi[prod]=int(product.iloc[a[0],4])
    Ri[prod]=int(product.iloc[a[0],5])
    Qi[prod]=float(product.iloc[a[0],10])
    Ki[prod]=int(product.iloc[a[0],6])


# add variables
Xit = m.addVars(P, T, vtype=GRB.INTEGER, name="Xit")  # purchase quantity
Yit = m.addVars(P, T, vtype=GRB.BINARY, name="Yit")  # whether buy or not
Nit = m.addVars(P, T, vtype=GRB.INTEGER, name="Nit")  # random int
Z = m.addVar(vtype=GRB.CONTINUOUS, name="Z")  # holding cost

# set objectives: min holding cost and purchase time
m.setObjective(gp.quicksum(Yit[i, t]*ordering_cost for i in P for t in range(5))+Z, GRB.MINIMIZE)

# add constraints

# 1. not purchase on weekends
for i in P:
    m.addConstr(Yit[i, 5] == 0)
    m.addConstr(Yit[i, 6] == 0)

# 2. purchase in unit
for i in P:
    for t in range(T):
        m.addConstr(Xit[i, t] == Nit[i, t] * Ki[i])

# 3. not exceed max inventory level
for i in P:
    for t in range(T):
        inventory = Si[i] + gp.quicksum(Xit[i, j] - Dit[i, j] for j in range(max(0, t-lead_time), t+1))
        m.addConstr(inventory <= SSpi[i])

# 4. not lower than SS
for i in P:
    for t in range(T):
        inventory = Si[i] + gp.quicksum(Xit[i, j] - Dit[i, j] for j in range(max(0, t-lead_time), t+1))
        m.addConstr(inventory >= Ri[i])

# 5. max holding cost
holding_cost = gp.quicksum((Si[i] + gp.quicksum(Xit[i, t-lead_time] - Dit[i, t] for t in range(T))) * Qi[i] for i in P)
m.addConstr(Z >= holding_cost)

# 6. if 'buy', we buy
for i in P:
    for t in range(T):
        m.addConstr(Xit[i, t] <= M * Yit[i, t])
        m.addConstr(Xit[i, t] >= 1 * Yit[i, t])

# 7. service level: <= $125, 95%; otherwise 100%
for i in P:
    for t in range(T):
        if Qi[i] > 125:
            m.addConstr(Xit[i, t] >= Dit[i, t])
        else:
            m.addConstr(Xit[i, t] >= 0.95 * Dit[i, t])

# 8. ensure we have enough stock on weekends
for i in P:
    m.addConstr(
        Si[i] + gp.quicksum(Xit[i, j] - Dit[i, j] for j in range(4)) >= (Dit[i, 5] + Dit[i, 6])
    )

# 9. prefer purchase 1~3 times per week
for i in P:
    m.addConstr(gp.quicksum(Yit[i, t] for t in range(5)) >= 1)
    m.addConstr(gp.quicksum(Yit[i, t] for t in range(5)) <= 3)

m.optimize()

# print purchase result and stock level change, and output to Optimization Sheet
result={'Product':[] , 'Day':[] , 'Qty':[]}
stock={'Product':[],
    'Monday':[],
    'Tuesday':[],
    'Wednesday':[],
    'Thursday':[],
    'Friday':[],
    'Saturday':[],
    'Sunday':[]
}

if m.status == GRB.OPTIMAL:
    for i in P:
        stock['Product'].append(str(i))
        for t in range(T):
            qty=Xit[i, t].x
            s_level=Si[i] + sum(Xit[i, tau].x - Dit[i, tau] for tau in range(t + 1))
            if t==0:
                stock['Monday'].append(int(s_level))
            elif t==1:
                stock['Tuesday'].append(int(s_level))
            elif t == 2:
                stock['Wednesday'].append(int(s_level))
            elif t == 3:
                stock['Thursday'].append(int(s_level))
            elif t == 4:
                stock['Friday'].append(int(s_level))
            elif t == 5:
                stock['Saturday'].append(int(s_level))
            elif t == 6:
                stock['Sunday'].append(int(s_level))

            if qty==0:
                continue
            else:
                print(f"Weekday: {t+1} | {i} | {Xit[i, t].x} unit")
                result['Product'].append(str(i))
                result['Day'].append(str(t))
                result['Qty'].append((str(qty)))

    print(f"MAX COST: {Z.x}")

else:
    print("No Result, ERROR!")

# ready in Excel, plot would be finished in Excel
with pd.ExcelWriter(stock_level, mode='a', engine='openpyxl', if_sheet_exists="replace") as writer:
    result.to_excel(writer, sheet_name='Purchase',index=False)
    stock.to_excel(writer, sheet_name='Inv Level', index=False)

