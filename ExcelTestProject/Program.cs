using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Collections;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using Floors;
using System.IO;
using System.Reflection;

namespace ExcelTestProject
{
  class Program
  {
    
    static void Main(string[] args)
    {
      //Console.WriteLine("Enter 1 for fifth floor OR 2 floor sixth floor");
      //int choice = Convert.ToInt32(Console.ReadLine());
      //var floor = GetFloorInstance(choice);
      //Console.WriteLine(floor.ToString());
      //floor.CheckExcellProcesses();
      //try
      //{
      //  //floor._ReadExcel = new Excel(@"D:\TFS SERVER\ExcelTestProject\Files\Employees_6thFloor.dat", 1);
      //  string path = @"D:\TFS SERVER\ExcelTestProject\Files\Employees_5thFloor.xls";
      //  floor._ReadExcel = new Excel(@path, 1);
      //  floor._SaveLocation = "EXCEL";
      //  floor.CreateMultipleFiles();
      //  floor._ReadExcel.Close();
      //  floor._writeExcel.Close();
      //}
      //catch (Exception e)
      //{
      //  Console.WriteLine(e.StackTrace);
      //  floor.KillExcel();
      //  Console.ReadKey();
      //}
      //floor.KillExcel();
      //Console.ReadKey();
      KillExcel();
    }

    public static FloorNo  GetFloorInstance(int choice)
    {
      if (choice == 1)
        return new FifthFloor();
      
      return new SixthFloor();
    }
    public static void KillExcel()
    {
      Process[] AllProcesses = Process.GetProcessesByName("excel");
      // check to kill the right process
      foreach (Process ExcelProcess in AllProcesses)
      {
       
          ExcelProcess.Kill();
      }
      AllProcesses = null;
    }

  }
}
