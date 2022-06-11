# -*- coding: utf-8 -*-
"""

@author:Case Study Group 6
"""



#Import puLP optimizer
import pulp
import pandas as pd

#loading building profile timeseries############################################################################################
Data = pd.read_excel('profile.xlsx', sep = ',', thousands = ',', index_col=0)

#creating Electric load dataframe#################################################################################################
EL_load = Data["el_load_house[kW]"]
EL_load = EL_load.to_frame().to_dict('index')

#creating Thermal/Heat load dataframe##################################################################################################
HL_load = Data["th_load_house[kW]"]
HL_load = HL_load.to_frame().to_dict('index')

#creating pv generation dataframe###########################################################################################
PV_generation = Data["PV[kW]"]
PV_generation = PV_generation.to_frame().to_dict('index')


#System Parameters (Dimensioning)
eff_storage=.98  #efficiency of storage (2% loss every 1 hour)
cop_pump=3      #COP of heat pump
eff_boiler=.98  #Efficiency of boiler
E_hp_max=20      # Maximum heat output from the Heat Pump
H_boiler_max=20 #Maximum heat output from Boiler
H_st_max=200    #Maximum Heat Storage Capacity
ap = 20     #installed photovoltaic installed capacity in Kilowatts
cost=0.25   #Electricty price in Euro/KWh
fit=0.10    #Feedin Tariff in Euro/KWh
SOC_0 = 100 #Assumed initialstate of Charge
dt = -1     #Storage control

#Declaring problem variables
El_grid = pulp.LpVariable.dicts("Grid Electricity", (k for k in Data.index), lowBound=0)
PV_feed = pulp.LpVariable.dicts("PV feed-in", (k for k in Data.index), lowBound=0)
PV_self = pulp.LpVariable.dicts("PV self", (k for k in Data.index), lowBound=0)
E_hp = pulp.LpVariable.dicts("Electrical_Power_HP", (k for k in Data.index), lowBound=0)
H_hp = pulp.LpVariable.dicts("Thermal_HP", (k for k in Data.index), lowBound=0)
E_boiler = pulp.LpVariable.dicts("Electrical_Power_Boiler", (k for k in Data.index), lowBound=0)
H_boiler = pulp.LpVariable.dicts("Thermal_Boiler", (k for k in Data.index), lowBound=0)
H_st = pulp.LpVariable.dicts("Thermal_Storage", (k for k in Data.index), lowBound=0)
SOC = pulp.LpVariable.dicts("State of Charging", (k for k in Data.index), lowBound=0)

# Create the 'prob' variable to contain the problem data##########################################################################
model = pulp.LpProblem("Operation Cost Minimization problem", pulp.LpMinimize)

# The objective function is added to 'model' first############################################################################
model += pulp.lpSum( 
        [ (El_grid[k] - PV_feed[k])for k in Data.index]
        )


#constraints
for k in Data.index: 
    model += El_grid[k] + PV_self[k] == EL_load[k]["el_load_house[kW]"] + E_hp[k] + E_boiler[k] 
    model += H_st[k] + H_hp[k] + H_boiler[k] == HL_load[k]["th_load_house[kW]"]      
    model += H_hp[k] == cop_pump * E_hp[k] 
    model += H_boiler[k] == eff_boiler * E_boiler[k]
    model += PV_self[k] + PV_feed[k] == PV_generation[k]['PV[kW]']
    model += 0 <= SOC[k] <= H_st_max
    model += 0<= E_hp[k] <= E_hp_max
    model += 0<= H_boiler[k] <= H_boiler_max
    model += 0<=El_grid[k]
    model += 0<=PV_feed[k]
    model += 0<=PV_self[k]

model += SOC[1] == SOC_0*eff_storage  #state of charge at the beginning of the year

for k in range (1,8758):
    
    if  PV_generation[k+1]['PV[kW]'] >= (EL_load[k+1]["el_load_house[kW]"] + E_boiler[k+1] + E_hp[k+1]):
            model += SOC[k+1] == SOC[k]*eff_storage - (H_st[k])*dt  #charging the storage
           
    else:
#        PV_generation[k] <= (EL[k] + E_boiler[k] + E_hp[k]):
            model += SOC[k+1] == SOC[k]*eff_storage - (H_st[k])*(-dt) #discharging the storage
    

# The problem data is written to an .lp file
model.writeLP('Grid Peak Minimization problem.lp')

# The problem is solved using PuLP's choice of Solver
model.solve()
pulp.LpStatus[model.status]


# Print our objective function value (Total Costs)
print (pulp.value(model.objective))

output = []
for k in Data.index:
    var_output = {
        'Time': k,
        'El_grid': El_grid[k].varValue,
        'PV_feed': PV_feed[k].varValue,
        'PV_self': PV_self[k].varValue,
        'E_hp':    E_hp[k].varValue,
        'H_hp':    H_hp[k].varValue,
        'E_boiler': E_boiler[k].varValue,
        'H_boiler': H_boiler[k].varValue,
    }
    output.append(var_output)
output_df = pd.DataFrame.from_records(output).sort_values(['Time'])
output_df.set_index(['Time'], inplace=True)
output_df

### Write Files
file_name = "Grid Peak Minimization problem.csv"
output_df.to_csv(file_name, sep=',', encoding='utf-8')

