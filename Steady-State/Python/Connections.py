# -*- coding: utf-8 -*-
"""
Created on Thu Dec 17 11:52:36 2020

@author: u0128100
"""
import numpy as np
import win32com.client as win32
import os
import time
 
parent = os.path.dirname(os.getcwd()) 

Inputs=np.asarray(np.genfromtxt(os.path.join(os.path.dirname(os.getcwd()) , 'Inputs.csv'), delimiter=','))
cor=[1,100,1/3600,1/3600,1/3600]

hysys = win32.Dispatch('hysys.Application')
hysys.visible=True


for iteration in range(1):
    
    #Direct connection
    simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\Sour Water Stripper.hsc"))
    simulation.Activate()
    
    stream=simulation.Flowsheet.streams["SourH2O Feed"]
    gasStream=simulation.Flowsheet.streams["Off Gas"]
    
    output_direct=np.zeros((Inputs.shape[0],Inputs.shape[1]+1))
    
    time_direct_real=0
    time_direct_flex=0
    for i in range(Inputs.shape[0]):
        
        simulation.Solver.canSolve=False
        start=time.time()
        stream.TemperatureValue=Inputs[i,0]
        time_direct_real=time_direct_real+(time.time()-start)
        simulation.Solver.canSolve=True
        output_direct[i,0]=stream.TemperatureValue
        
        simulation.Solver.canSolve=False
        start=time.time()
        stream.PressureValue=Inputs[i,1]*100
        time_direct_real=time_direct_real+(time.time()-start)
        simulation.Solver.canSolve=True
        output_direct[i,1]=stream.PressureValue
        
        simulation.Solver.canSolve=False
        start=time.time()
        temp=list(stream.ComponentMolarFlowValue)
        temp[0]=Inputs[i,2]/3600
        stream.ComponentMolarFlowValue=temp
        time_direct_flex=time_direct_flex+(time.time()-start)
        simulation.Solver.canSolve=True
        output_direct[i,2]=stream.ComponentMolarFlowValue[0]
        
        simulation.Solver.canSolve=False
        start=time.time()
        temp=list(stream.ComponentMolarFlowValue)
        temp[1]=Inputs[i,3]/3600
        stream.ComponentMolarFlowValue=temp
        temp=list(stream.ComponentMolarFlowValue)
        time_direct_flex=time_direct_flex+(time.time()-start)
        simulation.Solver.canSolve=True
        output_direct[i,3]=stream.ComponentMolarFlowValue[1]
        
        simulation.Solver.canSolve=False
        start=time.time()
        temp=list(stream.ComponentMolarFlowValue)
        temp[2]=Inputs[i,4]/3600
        stream.ComponentMolarFlowValue=temp 
        time_direct_flex=time_direct_flex+(time.time()-start)
        simulation.Solver.canSolve=True
        output_direct[i,4]=stream.ComponentMolarFlowValue[2]
        
        output_direct[i,5]=gasStream.ComponentMolarFlowValue[0]
        
        print(i)
    
    time_direct_real=time_direct_real/(i+1)/2
    time_direct_flex=time_direct_flex/(i+1)/3
    simulation.Close(SaveChanges=False)
    
    
    #Indirect Connection
    
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
    
    
    simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\Sour Water Stripper.hsc"))
    simulation.Activate()
    
    stream=simulation.Flowsheet.streams["SourH2O Feed"]
    gasStream=simulation.Flowsheet.streams["Off Gas"]
    
    InputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFlow,True,0),
            HysysValue(stream.ComponentMolarFlow,True,1),HysysValue(stream.ComponentMolarFlow,True,2)]
    
    OutputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFlow,True,0),
            HysysValue(stream.ComponentMolarFlow,True,1),HysysValue(stream.ComponentMolarFlow,True,2),
            HysysValue(gasStream.ComponentMolarFlow,True,0)]
    
    
    output_indirect=np.zeros((Inputs.shape[0],Inputs.shape[1]+1))
    time_indirect=0
    for i in range(Inputs.shape[0]):
        for j in range(Inputs.shape[1]):
            simulation.Solver.canSolve=False
            start=time.time()
            InputVars[j].value=Inputs[i,j]*cor[j]
            time_indirect=time_indirect+(time.time()-start)
            simulation.Solver.canSolve=True
            
            output_indirect[i,j]=OutputVars[j].value
        output_indirect[i,j+1]=OutputVars[j+1].value
        print(i)   
    time_indirect=time_indirect/(i+1)/j
    simulation.Close(SaveChanges=False)
    
    
    
    #Spreadsheet Connection
    #cor=[1,100,1/3600,1/3600,1/3600]
    simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\Sour Water Stripper_Spreadsheet.hsc"))
    simulation.Activate()
    
    sheet=simulation.Flowsheet.Operations["Sheet"];
    
    output_sheet=np.zeros((Inputs.shape[0],Inputs.shape[1]+1))
    time_sheet=0
    for i in range(Inputs.shape[0]):
        for j in range(Inputs.shape[1]):
            simulation.Solver.canSolve=False
            start=time.time()
            sheet.Cell(0,j).CellValue=Inputs[i,j]#*cor[j]
            time_sheet=time_sheet+(time.time()-start)
            simulation.Solver.canSolve=True
            
            output_sheet[i,j]=sheet.Cell(1,j).CellValue
        output_sheet[i,j+1]=sheet.Cell(1,j+1).CellValue
        print(i)   
    time_sheet=time_sheet/(i+1)/j
        
    simulation.Close(SaveChanges=False)

    
    np.savetxt('times' + '.csv',[time_direct_real,time_direct_flex,time_indirect,time_sheet],delimiter=',')
    np.savetxt('results_direct'+ '.csv',output_direct,delimiter=',')
    np.savetxt('results_indirect' + '.csv',output_indirect,delimiter=',')
    np.savetxt('results_sheet' + '.csv',output_sheet,delimiter=',')

#DataTable Connection

simulation = hysys.SimulationCases.Open(os.path.join(os.path.dirname(os.getcwd()), "Hysys\\Sour Water Stripper_DataTable.hsc"))
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
    

