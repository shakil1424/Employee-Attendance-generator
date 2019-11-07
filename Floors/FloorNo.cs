using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace Floors
{
    public abstract class FloorNo
    {
    public  static Hashtable myHashtable;
    public  string _currentEmployeeName = "";
    public  DateTime _initialDateTime = new DateTime();
    public  int _sourceRowCount = 0;
    public  int _sourceColumnCount = 0;
    public  Excel _ReadExcel = new Excel();
    public  Excel _writeExcel = new Excel();
    public  int _noOfDaysInMonth = 0;
    public  Excel _dummyExcel = new Excel();
    public  int _nameColumnNo;
    public  int _dateColumnNo;
    public  int _startingRowIndex;
    public string _SaveLocation { get; set; }

    public void InitializeExcelParameters()
    {
      _currentEmployeeName = _ReadExcel.ReadCell(_startingRowIndex, _nameColumnNo);
      _initialDateTime = GetDateTime(_ReadExcel.ReadCell(_startingRowIndex, _dateColumnNo));
      if (_sourceRowCount == 0 && _sourceColumnCount == 0)
      {
        _sourceRowCount = _ReadExcel.CountUsedRows();
        _sourceColumnCount = _ReadExcel.CountUsedColumns();
      }
      Console.WriteLine("Initialization Done ");
    }
    public  void CreateNewFile()
    {
      var executionDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
      var path = executionDirectory + "\\Resources\\NewResult.xlsx";
      _dummyExcel = new Excel(path, 1);
      _dummyExcel.WriteToCell(1, 3, _currentEmployeeName);
      _dummyExcel.WriteToCell(2, 3, _initialDateTime.ToString("yyyy"));
      _dummyExcel.WriteToCell(3, 3, _initialDateTime.ToString("MMMM"));
      for (int i = 1; i <= DateTime.DaysInMonth(_initialDateTime.Year, _initialDateTime.Month); i++)
      {
        DateTime dateTime = new DateTime(_initialDateTime.Year, _initialDateTime.Month, i);
        string day = dateTime.DayOfWeek.ToString();
        _dummyExcel.WriteToCell(6 + i, 1, day);
        _dummyExcel.WriteToCell(6 + i, 2, dateTime.ToString("dd/MMMM/yyyy"));
      }
      string saveFileName = _SaveLocation+@"\" + _currentEmployeeName + "_" + _initialDateTime.ToString("MMMM") + "_" + _initialDateTime.ToString("yyyy");
      _dummyExcel.SaveAs(@saveFileName);
      _dummyExcel.Close();
      _writeExcel = new Excel(@saveFileName, 1);
       Console.WriteLine("For employee " + _writeExcel.ReadCell(1, 3) + " New File Is Created");
    }

    public  DateTime GetDateTime(string dateTime)
    {
      var newDatetime = new DateTime();
      double date;
      if (Double.TryParse(dateTime, out date))
      {
        date = Double.Parse(dateTime);
        newDatetime = DateTime.FromOADate(date);
      }
      else
      {
        if (dateTime.Length > 8)
        {
          var dateParts = dateTime.Split('-');
          int day = Convert.ToInt32(dateParts[0]);
          int month = Convert.ToInt32(dateParts[1]);
          int year = Convert.ToInt32(dateParts[2]);
          newDatetime = new DateTime(year, month, day);
        }
        else
        {
          newDatetime = DateTime.Parse(dateTime);
        }
      }
      return newDatetime;
    }
    public void CheckExcellProcesses()
    {
      Process[] AllProcesses = Process.GetProcessesByName("excel");
      myHashtable = new Hashtable();
      int iCount = 0;

      foreach (Process ExcelProcess in AllProcesses)
      {
        myHashtable.Add(ExcelProcess.Id, iCount);
        iCount = iCount + 1;
      }
    }
    public void KillExcel()
    {
      Process[] AllProcesses = Process.GetProcessesByName("excel");
      // check to kill the right process
      foreach (Process ExcelProcess in AllProcesses)
      {
        if (myHashtable.ContainsKey(ExcelProcess.Id) == false)
          ExcelProcess.Kill();
      }
      AllProcesses = null;
    }
    public abstract void CreateMultipleFiles();
    }
}
