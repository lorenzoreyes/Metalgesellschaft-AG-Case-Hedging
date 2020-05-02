#1 import packages needed
import math
import numpy as np
import pandas as pd
from pandas import ExcelWriter
import xlsxwriter

#2 setting object binomial set of variables of metalgesellschaft
# Spot, Time, u, d, in this model sigma is constant
# C as the price you sold contracts
# M steps you run the simulation
# b quantity of barrels you are compromised to sold at To

S0 = 20.0
T = 10.0
u = 0.1
d = 0.1
C = 24.0
M = float
b = 1600.0 # in millions barrels compromised
delta_b = 160.0 # in millions, you have 10 deliveries

# Note: you start selling at C 24.0 with a S Spot of 20.0,
# the difference multiply by delta barresls give you a
# 640.0 of Margin Safety to operate

#3 function binolia procedure
def simulate_tree(M):
    u = 0.1
    d = 0.1
    S = np.zeros((M + 1, M + 1))
    S[0, 0] = S0
    z = 1
    for t in range(1, M + 1):
        for i in range(z):
            S[i, t] = S[i, t-1] * (1 + u)
            S[i+1, t] = S[i, t-1] * (1 - d)
        z += 1
    return S

# transform the tree of 10 steps into pandas dataframe
Sts = pd.DataFrame(data=simulate_tree(10)[0:,0:])

# get operative results as the evolution of your deliveries given the spot at t and apply delta barrels

operative_results = (C - Sts) * delta_b

operative_results = operative_results.replace(to_replace=3840.0, value=0.0)

# get financial results as the value of the contracts plus the impact of deliveries
# here you start again the tree applying the barrels contracts in addition with deliveries
# as a result you get the financial result, but as you get the evolution of the nodes
# you have the evolution of your initial budget of each fluctuation

financial_results = (b * Sts) + operative_results

writer = pd.ExcelWriter('metalgesellschaft case.xlsx', engine='xlsxwriter')

Sts.to_excel(writer, sheet_name='spot binomial tree')

operative_results.to_excel(writer, sheet_name='operative results deliveries')

financial_results.to_excel(writer, sheet_name='financial results remaining')

writer.save()













