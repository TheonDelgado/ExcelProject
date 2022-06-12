using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using ExcelProject;

public class Program
{
    static void Main(string[] args)
    {
        foreach(object thing in ExcelWorker.data)
        {
            Console.WriteLine(thing);
        }
        EndProgram();
    }

    private static void EndProgram()
    {
        Console.WriteLine("Press any key to end program...");
        Console.ReadKey();
    }
}