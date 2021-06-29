# -*- coding: utf-8 -*-
"""
Created on Thu Dec 17 11:52:36 2020

@author: u0128100
"""
import numpy as np
import win32com.client as win32
import os
import time
import matplotlib.pyplot as plt
#from decimal import Decimal    
 
parent = os.path.dirname(os.getcwd()) 

Inputs=np.asarray(np.genfromtxt(os.path.join(os.path.dirname(os.getcwd()) , 'Inputs.csv'), delimiter=','))
Inputs=Inputs[:,:]

hysys = win32.Dispatch('hysys.Application')
hysys.visible=True

timeStep=60


#Direct connection
simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\De-Propanizer - Dynamic Model.hsc"))
simulation.Activate()

stream=simulation.Flowsheet.streams["Tower Feed"]
reacStream=simulation.Flowsheet.streams["Bttms-1"]

output_direct=np.zeros((Inputs.shape[0],Inputs.shape[1]+1))


time_direct_real=0
time_direct_flex=0
for i in range(Inputs.shape[0]):
    
    simulation.Solver.Integrator.IsRunning = False;    
    start=time.time()
    stream.TemperatureValue=Inputs[i,0]
    time_direct_real=time_direct_real+(time.time()-start)
    output_direct[i,0]=stream.TemperatureValue
    
    start=time.time()
    stream.PressureValue=Inputs[i,1]
 #   stream.MolarFlowValue=Inputs[i,1]/3600
    time_direct_real=time_direct_real+(time.time()-start)
    output_direct[i,1]=stream.PressureValue
    
    start=time.time()
    stream.ComponentMolarFractionValue=Inputs[i,2:]
    time_direct_flex=time_direct_flex+(time.time()-start)
    
    output_direct[i,2:-1]=stream.ComponentMolarFractionValue
    
    
    simulation.Solver.Integrator.StopTime= simulation.Solver.Integrator.CurrentTimeValue+timeStep;
    simulation.Solver.Integrator.IsRunning = True;    
    
    while(simulation.Solver.Integrator.IsRunning):
       time.sleep(0.01) 
    
    
    output_direct[i,-1]=reacStream.ComponentMolarFractionValue[2]
    print(i)

time_direct_real=time_direct_real/(i+1)/2
time_direct_flex=time_direct_flex/(i+1)
simulation.Close(SaveChanges=False)
hysys.Quit()

time.sleep(10)
#Indirect Connection
hysys = win32.Dispatch('hysys.Application')
hysys.visible=True

class HysysValue:
    def __init__(self,obj,isFlexVar=False, index=-1):
        self.obj=obj
        self.isFlexVar=isFlexVar
        self.index=index
    @property
    def value(self):
        if self.isFlexVar:
            if self.index==-1:
                return self.obj.Values
            else:
                return self.obj.Values[self.index]
        else:
            return self.obj.Value
    @value.setter
    def value(self, newVal):
        if self.isFlexVar:
            if self.index==-1:
                self.obj.Values=newVal
            else:
                temp=list(self.obj.Values)
                temp[self.index]=newVal
                self.obj.Values=temp
        else:
            self.obj.Value =newVal


simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\De-Propanizer - Dynamic Model.hsc"))
simulation.Activate()

stream=simulation.Flowsheet.streams["Tower Feed"]
reacStream=simulation.Flowsheet.streams["Bttms-1"]

InputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFraction,True,-1)]

OutputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFraction,True,-1),
        HysysValue(reacStream.ComponentMolarFraction,True,2)]


output_indirect=np.zeros((Inputs.shape[0],Inputs.shape[1]+1))
time_indirect=0
for i in range(Inputs.shape[0]):
    simulation.Solver.canSolve=False
        
    start=time.time()
    InputVars[0].value=Inputs[i,0]
    time_indirect=time_indirect+(time.time()-start)
    
    output_indirect[i,0]=OutputVars[0].value
    
    start=time.time()
    InputVars[1].value=Inputs[i,1]#/3600
    time_indirect=time_indirect+(time.time()-start)
    
    output_indirect[i,1]=OutputVars[1].value
    
    start=time.time()
    InputVars[2].value=Inputs[i,2:]
    time_indirect=time_indirect+(time.time()-start)
    
    output_indirect[i,2:-1]=OutputVars[2].value
    
    simulation.Solver.Integrator.StopTime= simulation.Solver.Integrator.CurrentTimeValue+timeStep;
    simulation.Solver.Integrator.IsRunning = True;    
    
    while(simulation.Solver.Integrator.IsRunning):
       time.sleep(0.01)     
    
    output_indirect[i,-1]=OutputVars[3].value
    
    print(i)   
time_indirect=time_indirect/(i+1)/3
simulation.Close(SaveChanges=False)
hysys.Quit()

time.sleep(10)
#Spreadsheet Connection
hysys = win32.Dispatch('hysys.Application')
hysys.visible=True
simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\De-Propanizer - Dynamic Model_Spreadsheet.hsc"))
simulation.Activate()

sheet=simulation.Flowsheet.Operations["Sheet"];

stream=simulation.Flowsheet.streams["Tower Feed"]

output_sheet=np.zeros((Inputs.shape[0],Inputs.shape[1]+1))
time_sheet=0
for i in range(Inputs.shape[0]):
    for j in range(Inputs.shape[1]):
        simulation.Solver.canSolve=False
        start=time.time()
        sheet.Cell(0,j).CellValue=Inputs[i,j]#*cor[j]
        time_sheet=time_sheet+(time.time()-start) 
        
        output_sheet[i,j]=sheet.Cell(1,j).CellValue
        
    simulation.Solver.canSolve=True
    simulation.Solver.Integrator.IsRunning = True;  
    simulation.Solver.canSolve=True
    simulation.Solver.Integrator.IsRunning = True;
    
    simulation.Solver.Integrator.IsRunning = True;  
    simulation.Solver.Integrator.StopTime= simulation.Solver.Integrator.CurrentTimeValue+timeStep;
    simulation.Solver.Integrator.IsRunning = True;    
    
    time.sleep(1)   
    while(simulation.Solver.Integrator.IsRunning):
       time.sleep(1)   
    output_sheet[i,-1]=sheet.Cell(1,output_sheet.shape[1]-1).CellValue
    print(i)   
time_sheet=time_sheet/(i+1)/j
    
simulation.Close(SaveChanges=False)
hysys.Quit()

i=-1
#plt.plot(Inputs[:100,i])
plt.plot(output_direct[:100,i])
plt.plot(output_indirect[:100,i])
plt.plot(output_sheet[:100,i])

np.savetxt('times.csv',[time_direct_real,time_direct_flex,time_indirect,time_sheet],delimiter=',')
np.savetxt('results_direct.csv',output_direct,delimiter=',')
np.savetxt('results_indirect.csv',output_indirect,delimiter=',')
np.savetxt('results_sheet.csv',output_sheet,delimiter=',')

#DataTable Connection

simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\De-Propanizer - Dynamic Model_DataTable.hsc"))
simulation.Activate()

table=simulation.DataTables["DataTable1"]
table.StartTransfer()
Wtags=list(table.WriteTags)
Rtags=list(table.ReadTags)

output_dataTable=np.zeros((Inputs.shape[0],Inputs.shape[1]+1))
time_dataTable=0
for i in range(Inputs.shape[0]):
    for j in range(Inputs.shape[1]):
        start=time.time()
        table.DataValue=Inputs[i,j]#*cor[j]
        time_dataTable=time_dataTable+(time.time()-start)
        
    print(i)   
    print(time.time()-start) 
    

