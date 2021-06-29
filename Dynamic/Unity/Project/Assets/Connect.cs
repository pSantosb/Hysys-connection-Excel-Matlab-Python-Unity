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

        double TimeStep = 60;

        List<List<double>> Inputs = Readcsv(filePath, separator);
        Debug.Log(Inputs.Count);

        HYSYS.Application HyApp = new HYSYS.Application();
        HyApp.Visible = true;

        HYSYS.SimulationCase simulation; HYSYS.ProcessStream stream; HYSYS.ProcessStream reacStream;

        double[] cor = new double[] { 1, 100, 1 / 3600, 1 / 3600, 1 / 3600 };
        double start;
        //Direct connection

        filePath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName + Path.DirectorySeparatorChar + "Hysys" + Path.DirectorySeparatorChar + "De-Propanizer - Dynamic Model.hsc";
        Debug.Log(filePath);
        simulation = (HYSYS.SimulationCase)HyApp.SimulationCases.Open(filePath);
        simulation.Activate();


        stream = (HYSYS.ProcessStream)simulation.Flowsheet.Streams["Tower Feed"];
        reacStream = (HYSYS.ProcessStream)simulation.Flowsheet.Streams["Bttms-1"];
        List<double> temp;
        Debug.Log(stream.name);

        double time_direct_real = 0;
        List<List<double>> output_direct = new List<List<double>>();

        for (int i = 0; i < Inputs.Count; i++)
        {
            output_direct.Add(new List<double>());

            simulation.Solver.CanSolve = false;
            sw.Start();
            stream.TemperatureValue = Inputs[i][0];
            sw.Stop();
            time_direct_real = time_direct_real + sw.ElapsedMilliseconds;
            sw.Reset();
            simulation.Solver.CanSolve = true;
            output_direct[i].Add(stream.TemperatureValue);

            simulation.Solver.CanSolve = false;
            sw.Start();
            stream.PressureValue = Inputs[i][1] * 100;
            sw.Stop();
            time_direct_real = time_direct_real + sw.ElapsedMilliseconds;
            sw.Reset();
            simulation.Solver.CanSolve = true;
            output_direct[i].Add(stream.PressureValue);

            temp = (List<double>)stream.ComponentMolarFractionValue;//Property cannot be accessed
            Debug.Log(temp);

            for (int j = 2; j < Inputs[i].Count; j++)
            {
                output_direct[i].Add(0);
            }

            simulation.Solver.Integrator.StopTime = simulation.Solver.Integrator.CurrentTimeValue + TimeStep;
            simulation.Solver.Integrator.IsRunning = true;
            while (simulation.Solver.Integrator.IsRunning) { System.Threading.Thread.Sleep(1000); }

        }
        time_direct_real = time_direct_real / Inputs.Count / 2;

        makecsv(output_direct, Directory.GetParent(Directory.GetCurrentDirectory()).FullName + Path.DirectorySeparatorChar + "results_direct.csv");

 //         simulation.Close();
 //         HyApp.Quit();

        // Indirect Connection
 //      HyApp = new HYSYS.Application();
 //       HyApp.Visible = true;

        temp = (List<double>)stream.ComponentMolarFraction.Values; //Property cannot be accessed
        Debug.Log(temp);

   //     simulation.Close();
   //     HyApp.Quit();
        System.Threading.Thread.Sleep(1000);

        // Spreadsheet Connection
      //  HyApp = new HYSYS.Application();
      ///  HyApp.Visible = true;

        filePath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).FullName).FullName + Path.DirectorySeparatorChar + "Hysys" + Path.DirectorySeparatorChar + "De-Propanizer - Dynamic Model_Spreadsheet.hsc";
        simulation = (HYSYS.SimulationCase)HyApp.SimulationCases.Open(filePath);
        simulation.Activate();

        List<List<double>> output_sheet = new List<List<double>>();

        sheet = (HYSYS.SpreadsheetOp)simulation.Flowsheet.Operations[null]["Sheet"];
        Debug.Log(sheet.name);

        HYSYS.SpreadsheetCell cell;
        double time_sheet = 0;

        for (int i = 0; i < Inputs.Count; i++)
        {
            output_sheet.Add(new List<double>());
            for (int j = 0; j < Inputs[i].Count; j++)
            {
                simulation.Solver.CanSolve = false;
                cell = sheet.Cell[0, j];
                sw.Start();
                cell.CellValue = Inputs[i][j];//#*cor[j]
                sw.Stop();
                time_sheet = time_sheet + sw.ElapsedMilliseconds;
                sw.Reset();
                simulation.Solver.CanSolve = true;
                output_sheet[i].Add(sheet.Cell[1, j].CellValue);
            }

            simulation.Solver.Integrator.StopTime = simulation.Solver.Integrator.CurrentTimeValue + TimeStep;
            simulation.Solver.Integrator.IsRunning = true;
            while (simulation.Solver.Integrator.IsRunning) { System.Threading.Thread.Sleep(1000); }
            
            output_sheet[i].Add(sheet.Cell[1, Inputs[i].Count].CellValue);
        }
        time_sheet = time_sheet / Inputs.Count / Inputs[0].Count;
        //simulation.Close();
        //   HyApp.Quit();
        makecsv(output_sheet, Directory.GetParent(Directory.GetCurrentDirectory()).FullName + Path.DirectorySeparatorChar + "results_sheet.csv");

        List<List<double>> times = new List<List<double>>() { new List<double>() { time_direct_real, 0, 0, time_sheet, 0 } };
        Debug.Log(times);
        makecsv(times, Directory.GetParent(Directory.GetCurrentDirectory()).FullName + Path.DirectorySeparatorChar + "times.csv");
      //  simulation.Close();
      //  HyApp.Quit();

   
    }

    // Update is called once per frame
    void Update()
    {
     
    }

    List<List<double>> Readcsv(string filePath, char separator)
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

                for (int j = 0; j < values.Length; j++)
                {

                    Inputs[i].Add(System.Convert.ToDouble(values[j].Replace(".", ",")));

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
                    if (j != data[i].Count - 1)
                    {
                        finalString = finalString + ",";
                    }
                }

                sw.WriteLine(finalString);
                finalString = "";
            }
        }

    }
}

