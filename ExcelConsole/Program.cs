using ExcelProcessor.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"D:\Javra\Poc\EcelWork\ExcelProcessing\WebApplication1\ExcelData\Uploads\TS-4040_30972_Air_France_Added_Large.xlsx";
           // var path = @"C:\Users\pjha.JAVRA\Downloads\Sample - Superstore.xls";
            Stopwatch stopwatch = new Stopwatch();
            //OLEDB
            stopwatch.Start();
            var dt = XslLibrary.ReadDataTable(path);
            stopwatch.Stop();
            Console.WriteLine(stopwatch.ElapsedMilliseconds);
            stopwatch.Reset();

            //Interloop
            //stopwatch.Start();
            //XslLibrary.ExcelInterloop(path);
            //stopwatch.Stop();
            //Console.WriteLine(stopwatch.ElapsedMilliseconds);
            //stopwatch.Reset();

            //stopwatch.Start();
            //XslLibrary.WriteDatatableToExcel(dt);
            //stopwatch.Stop();
            //Console.WriteLine(stopwatch.ElapsedMilliseconds);

            stopwatch.Start();
            XslLibrary.WriteDatatableToDataset(dt);
            stopwatch.Stop();
            Console.WriteLine(stopwatch.ElapsedMilliseconds);

            //Console.WriteLine(stopwatch.ElapsedMilliseconds);
            //stopwatch.Start();
            //XslLibrary.ImportFromExcelNpoi(path);
            //stopwatch.Stop();
            //Console.WriteLine(stopwatch.ElapsedMilliseconds);
            //stopwatch.Reset();

            //Console.WriteLine(stopwatch.ElapsedMilliseconds);
            //stopwatch.Start();
            //XslLibrary.ReadOpenExcel(path);
            //stopwatch.Stop();
            //Console.WriteLine(stopwatch.ElapsedMilliseconds);
            //stopwatch.Reset();

            //Console.WriteLine(stopwatch.ElapsedTicks);
            //stopwatch.Start();
            //XslLibrary.XslReader(path);
            //stopwatch.Stop();
            //Console.WriteLine(stopwatch.ElapsedTicks);

            //Console.WriteLine(performance.RawValue);
            //var a = performance.NextValue();

            //Console.WriteLine(performance.RawValue);
            //var b = performance.NextValue();
            ////XslLibrary.ImportFromExcelNpoi(path);
            //Console.WriteLine(performance.RawValue);
            //var c = performance.NextValue();
            Console.ReadKey();
        }
    }
}
