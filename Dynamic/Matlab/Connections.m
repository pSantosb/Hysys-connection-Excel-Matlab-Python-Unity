


Inputs = table2array(readtable(fullfile(pwd,'..','Inputs.csv')));

hysys = actxserver('Hysys.Application');
hysys.visible=true;


timeStep=60;

%Direct connection
filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','De-Propanizer - Dynamic Model.hsc');
simulation=hysys.SimulationCases.Open(filePath{1});
simulation.Activate();

stream=simulation.Flowsheet.streams.Item("Tower Feed");
reacStream=simulation.Flowsheet.streams.Item("Bttms-1");

output_direct=zeros(size(Inputs,1),size(Inputs,2)+1);

time_direct_real=0;
time_direct_flex=0;
for i=1:1:size(Inputs,1)
    
    simulation.Solver.canSolve=false;
    tic;
    stream.TemperatureValue=Inputs(i,1);
    time_direct_real=time_direct_real+toc;
    output_direct(i,1)=stream.TemperatureValue;
    
    tic;
    stream.PressureValue=Inputs(i,2);
    time_direct_real=time_direct_real+toc;
    output_direct(i,2)=stream.PressureValue;
    
    tic;
%    stream.ComponentMolarFractionValue=Inputs(i,3:7);
    time_direct_flex=time_direct_flex+toc;
    output_direct(i,3:size(Inputs,2))=stream.ComponentMolarFractionValue;
    
   
    simulation.Solver.Integrator.StopTime= simulation.Solver.Integrator.CurrentTimeValue+timeStep;
    simulation.Solver.Integrator.IsRunning = true;
    output_direct(i,size(output_direct,2))=reacStream.ComponentMolarFractionValue(3);
    
    while(simulation.Solver.Integrator.IsRunning)
       pause(0.01) 
    end
    i
end
time_direct_real=time_direct_real/(i+1)/2;
time_direct_flex=time_direct_flex/(i+1);
simulation.Close(false)
hysys.Quit()
pause(1);

%Indirect connection
hysys = actxserver('Hysys.Application');
hysys.visible=true;
filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','De-Propanizer - Dynamic Model.hsc');
simulation=hysys.SimulationCases.Open(filePath{1});
simulation.Activate();

stream=simulation.Flowsheet.streams.Item("Tower Feed");
reacStream=simulation.Flowsheet.streams.Item("Bttms-1");

InputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFraction,true,-1)];

OutputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFraction,true,-1), ...
        HysysValue(reacStream.ComponentMolarFraction,true,3)];

output_indirect=zeros(size(Inputs,1),size(Inputs,2)+1);
time_indirect=0;
for i=1:1:size(Inputs,1)
    simulation.Solver.canSolve=false;
    for j=1:1:size(InputVars,2)
        if ~InputVars(j).isFlexVar
            tic
            InputVars(j).value=Inputs(i,j);
            time_indirect=time_indirect+toc;
            output_indirect(i,j)=OutputVars(j).value;
        end
    end
    simulation.Solver.Integrator.StopTime= simulation.Solver.Integrator.CurrentTimeValue+timeStep;
    simulation.Solver.Integrator.IsRunning = true;
    while(simulation.Solver.Integrator.IsRunning)
       pause(0.01) 
    end
    output_indirect(i,end)=OutputVars(4).value;
    i 
end
time_indirect=time_indirect/(i+1)/2;
simulation.Close(false)
hysys.Quit()
pause(1);

%Spreadsheet connection
hysys = actxserver('Hysys.Application');
hysys.visible=true;
filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','De-Propanizer - Dynamic Model_Spreadsheet.hsc');
simulation=hysys.SimulationCases.Open(filePath{1});
simulation.Activate();

sheet=simulation.Flowsheet.Operations.Item("Sheet");

output_sheet=zeros(size(Inputs,1),size(Inputs,2)+1);
time_sheet=0;
for i=1:1:size(Inputs,1)
    simulation.Solver.canSolve=false;
    for j=1:1:size(Inputs,2)
        t="A"+j;
        tic;
        sheet.Cell(t).CellValue=Inputs(i,j);
        time_sheet=time_sheet+toc;
        
        output_sheet(i,j)=sheet.Cell("B"+j).CellValue;
    end
    simulation.Solver.Integrator.StopTime= simulation.Solver.Integrator.CurrentTimeValue+timeStep;
    simulation.Solver.Integrator.IsRunning = true;
    while(simulation.Solver.Integrator.IsRunning)
       pause(1) 
    end
    output_sheet(i,j+1)=sheet.Cell("B"+(j+1)).CellValue;
    i
end
time_sheet=time_sheet/(i+1)/j;
    
simulation.Close(false)
hysys.Quit()
pause(1);


csvwrite('times.csv',[time_direct_real,time_direct_flex,time_indirect,time_sheet])
csvwrite('results_direct.csv',output_direct)
csvwrite('results_indirect.csv',output_indirect)
csvwrite('results_sheet.csv',output_sheet)


%Spreadsheet connection
filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','Propylene Glycol Production - Dynamic Model_DataTable.hsc');
simulation=hysys.SimulationCases.Open(filePath{1});
simulation.Activate();

DataTable1=simulation.DataTables.Item("DataTable1")
DataTable1.StartTransfer()
DataTable1.GetValues(DataTable1.WriteTags)
