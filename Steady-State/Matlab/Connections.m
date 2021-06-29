


Inputs = table2array(readtable(fullfile(pwd,'..','Inputs.csv')));
cor=[1,100,1/3600,1/3600,1/3600];

hysys = actxserver('Hysys.Application');
hysys.visible=true;


for iteration=1:1:1

    %Direct connection
    filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','Sour Water Stripper.hsc');
    simulation=hysys.SimulationCases.Open(filePath{1});
    simulation.Activate();

    stream=simulation.Flowsheet.streams.Item("SourH2O Feed");
    gasStream=simulation.Flowsheet.streams.Item("Off Gas");

    output_direct=zeros(size(Inputs,1),size(Inputs,2)+1);

    time_direct_real=0;
    time_direct_flex=0;
    for i=1:1:size(Inputs,1)

        simulation.Solver.canSolve=false;
        tic;
        stream.TemperatureValue=Inputs(i,1);
        time_direct_real=time_direct_real+toc;
        simulation.Solver.canSolve=true;
        output_direct(i,1)=stream.TemperatureValue;

        simulation.Solver.canSolve=false;
        tic;
        stream.PressureValue=Inputs(i,2)*100;
        time_direct_real=time_direct_real+toc;
        simulation.Solver.canSolve=true;
        output_direct(i,2)=stream.PressureValue;

        simulation.Solver.canSolve=false;
        tic;
        temp=stream.ComponentMolarFlowValue;
        temp(1)=Inputs(i,3)/3600;
    %    stream.ComponentMolarFlowValue=temp;
        time_direct_flex=time_direct_flex+toc;
        simulation.Solver.canSolve=true;
        output_direct(i,3)=stream.ComponentMolarFlowValue(1);

        simulation.Solver.canSolve=false;
        tic;
        temp=stream.ComponentMolarFlowValue;
        temp(2)=Inputs(i,4)/3600;
    %    stream.ComponentMolarFlowValue=temp;
        time_direct_flex=time_direct_flex+toc;
        simulation.Solver.canSolve=true;
        output_direct(i,4)=stream.ComponentMolarFlowValue(2);

        simulation.Solver.canSolve=false;
        tic;
        temp=stream.ComponentMolarFlowValue;
        temp(3)=Inputs(i,5)/3600;
    %    stream.ComponentMolarFlowValue=temp;
        time_direct_flex=time_direct_flex+toc;
        simulation.Solver.canSolve=true;
        output_direct(i,5)=stream.ComponentMolarFlowValue(3);

        output_direct(i,6)=gasStream.ComponentMolarFlowValue(2);

        i
    end
    time_direct_real=time_direct_real/(i+1)/2;
    time_direct_flex=time_direct_flex/(i+1)/3;
    simulation.Close(false)

    %Indirect connection
    filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','Sour Water Stripper.hsc');
    simulation=hysys.SimulationCases.Open(filePath{1});
    simulation.Activate();

    stream=simulation.Flowsheet.streams.Item("SourH2O Feed");
    gasStream=simulation.Flowsheet.streams.Item("Off Gas");

    InputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFlow,true,1), ...
            HysysValue(stream.ComponentMolarFlow,true,2),HysysValue(stream.ComponentMolarFlow,true,3)];

    OutputVars=[ HysysValue(stream.Temperature), HysysValue(stream.Pressure),HysysValue(stream.ComponentMolarFlow,true,1), ...
            HysysValue(stream.ComponentMolarFlow,true,2),HysysValue(stream.ComponentMolarFlow,true,3), ...
            HysysValue(gasStream.ComponentMolarFlow,true,2)];

    output_indirect=zeros(size(Inputs,1),size(Inputs,2)+1);
    time_indirect=0;
    for i=1:1:size(Inputs,1)
        for j=1:1:size(Inputs,2)
            if ~InputVars(j).isFlexVar
                simulation.Solver.canSolve=false;
                tic
                InputVars(j).value=Inputs(i,j)*cor(j);
                time_indirect=time_indirect+toc;
                simulation.Solver.canSolve=true;
            end
            output_indirect(i,j)=OutputVars(j).value;
        end
        output_indirect(i,j+1)=OutputVars(j+1).value;
        i 
    end
    time_indirect=time_indirect/(i+1)/2;
    simulation.Close(false)

    %Spreadsheet connection
    filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','Sour Water Stripper_Spreadsheet.hsc');
    simulation=hysys.SimulationCases.Open(filePath{1});
    simulation.Activate();

    sheet=simulation.Flowsheet.Operations.Item("Sheet");

    output_sheet=zeros(size(Inputs,1),size(Inputs,2)+1);
    time_sheet=0;
    for i=1:1:size(Inputs,1)
        for j=1:1:size(Inputs,2)
            simulation.Solver.canSolve=false;
            t="A"+j;
            tic;
            sheet.Cell(t).CellValue=Inputs(i,j);
            time_sheet=time_sheet+toc;
            simulation.Solver.canSolve=true;

            output_sheet(i,j)=sheet.Cell("B"+j).CellValue;
        end
        output_sheet(i,j+1)=sheet.Cell("B"+(j+1)).CellValue;
        i
    end
    time_sheet=time_sheet/(i+1)/j;

    simulation.Close(false)


    csvwrite("times" + ".csv",[time_direct_real,time_direct_flex,time_indirect,time_sheet])
    csvwrite("results_direct" + ".csv",output_direct)
    csvwrite("results_indirect" + ".csv",output_indirect)
    csvwrite("results_sheet" + ".csv",output_sheet)

end

%data Table connection
filePath=fullfile(getfield( fliplr(regexp(fileparts(pwd),'/','split')), {1} ),'Hysys','Sour Water Stripper_DataTable.hsc');
simulation=hysys.SimulationCases.Open(filePath{1});
simulation.Activate();

DataTable1=simulation.DataTables.Item("DataTable1")
DataTable1.StartTransfer()
DataTable1.GetValues(DataTable1.WriteTags)