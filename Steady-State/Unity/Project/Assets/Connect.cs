using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System.IO;

public class Connect : MonoBehaviour
{
    HYSYS.SpreadsheetOp sheet;
    // Start is called before the first frame update
    void Start()
    {
        System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
        string filePath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName + Path.DirectorySeparatorChar + "Inputs.csv"; 
        char separator = ',';

        List<List<double>> Inputs = Readcsv(filePath, separator);
        Debug.Log(Inputs.Count);

        HYSYS.Application HyApp = new HYSYS.Application();
        HyApp.Visible = true;

        HYSYS.SimulationCase simulation; HYSYS.ProcessStream stream; HYSYS.ProcessStream gasStream;

        double[] cor = new double[] { 1, 100, 1 / 3600, 1 / 3600, 1 / 3600 };
        double start;

        //Direct connection
        
        filePath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName + Path.DirectorySeparatorChar + "Hysys" + Path.DirectorySeparatorChar + "Sour Water Stripper.hsc";
        Debug.Log(filePath);
        simulation = (HYSYS.SimulationCase)HyApp.SimulationCases.Open(filePath);
        simulation.Activate();


        stream = (HYSYS.ProcessStream)simulation.Flowsheet.Streams["SourH2O Feed"];
        gasStream = (HYSYS.ProcessStream)simulation.Flowsheet.Streams["Off Gas"];
        List<double> temp;

        double time_direct_real = 0;
        List<List<double>> output_direct = new List<List<double>>();

        for (int i = 0; i < Inputs.Count; i++)
        {
            output_direct.Add(new List<double>());

            sw.Start();
            simulation.Solver.CanSolve = false;
            stream.TemperatureValue = Inputs[i][0];
            sw.Stop();
            time_direct_real = time_direct_real + sw.ElapsedMilliseconds;
            sw.Reset();
            simulation.Solver.CanSolve = true;
            output_direct[i].Add(stream.TemperatureValue);

            sw.Start();
            simulation.Solver.CanSolve = false;
            stream.PressureValue = Inputs[i][1] * 100;
            sw.Stop();
            time_direct_real = time_direct_real + sw.ElapsedMilliseconds;
            sw.Reset();
            simulation.Solver.CanSolve = true;
            output_direct[i].Add(stream.PressureValue);

            temp = (List<double>)stream.ComponentMolarFlowValue;//Proerty cannot be accessed
            Debug.Log(temp);

            output_direct[i].Add(0);
            output_direct[i].Add(0);
            output_direct[i].Add(0);
            output_direct[i].Add(0);
        }
        time_direct_real = time_direct_real / Inputs.Count / 2;

        makecsv(output_direct, Directory.GetParent(Directory.GetCurrentDirectory()).FullName + Path.DirectorySeparatorChar + "results_direct.csv");

        //  simulation.Close();
        //  HyApp.Quit();

        // Indirect Connection

        temp = (List<double>)stream.ComponentMolarFlow.Values; //Property cannot be accessed
        Debug.Log(temp);
        
        // Spreadsheet Connection

        filePath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName + Path.DirectorySeparatorChar + "Hysys" + Path.DirectorySeparatorChar + "Sour Water Stripper_Spreadsheet.hsc";
        simulation = (HYSYS.SimulationCase)HyApp.SimulationCases.Open(filePath);
        simulation.Activate();

        List<List<double>> output_sheet = new List<List<double>>();

        sheet = (HYSYS.SpreadsheetOp) simulation.Flowsheet.Operations[null]["Sheet"];
        Debug.Log(sheet.name);

        HYSYS.SpreadsheetCell cell;
        double time_sheet=0;

        for (int i = 0; i < Inputs.Count; i++)
        {
            output_sheet.Add(new List<double>());
            for(int j=0;j<Inputs[i].Count;j++)
            {
                sw.Start();
                simulation.Solver.CanSolve = false;
                cell = sheet.Cell[0, j];
                start = Time.time;
                cell.CellValue = Inputs[i][j];//#*cor[j]
                sw.Stop();
                time_sheet = time_sheet + sw.ElapsedMilliseconds;
                sw.Reset();
                simulation.Solver.CanSolve = true;
                output_sheet[i].Add(sheet.Cell[1, j].CellValue);
            }
            output_sheet[i].Add(sheet.Cell[1, Inputs[i].Count].CellValue);
        }
        makecsv(output_sheet, Directory.GetParent(Directory.GetCurrentDirectory()).FullName + Path.DirectorySeparatorChar + "results_sheet.csv");
        time_sheet = time_sheet / Inputs.Count / Inputs[0].Count;

        List<List<double>> times = new List<List<double>>() { new List<double>() { time_direct_real, 0, 0, time_sheet, 0 } };
        Debug.Log(times);
        makecsv(times, Directory.GetParent(Directory.GetCurrentDirectory()).FullName + Path.DirectorySeparatorChar + "times.csv");



    }

    // Update is called once per frame
    void Update()
    {

 }

    List<List<double>> Readcsv( string filePath, char separator)
    {
        List<List<double>> Inputs = new List<List<double>>();
        Debug.Log("Running");
        using (var reader = new StreamReader(filePath))
        {
            int i = 0;
            while (!reader.EndOfStream)
            {
                Inputs.Add(new List<double>());
                var line = reader.ReadLine();
                var values = line.Split(separator);

                for(int j=0;j<values.Length; j++)
                {

                    Inputs[i].Add(System.Convert.ToDouble(values[j].Replace(".",",")));

                }
                i++;
            }
        }
        return Inputs;
    }
    void makecsv(List<List<double>> data, string filePath)
    {
        string finalString = "";

        File.WriteAllText(filePath, string.Empty);
        using (StreamWriter sw = File.AppendText(filePath))
        {
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    finalString = finalString + data[i][j].ToString("E5").Replace(",", ".");
                    if (j!= data[i].Count-1)
                    {
                        finalString = finalString + ",";
                    }
                }
                //finalString = finalString + "\n";

                sw.WriteLine(finalString);
                finalString = "";
            }
        }

    }
}

