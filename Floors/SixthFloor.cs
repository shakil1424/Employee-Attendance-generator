using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Floors
{
  public class SixthFloor: FloorNo
  {
    public override void CreateMultipleFiles()
    {
      _startingRowIndex = 1;
      _nameColumnNo = 2;
      _dateColumnNo = 3;
      InitializeExcelParameters();
      CreateNewFile();
      int targetIndex = 1;
      int sourceIndex = _startingRowIndex;
      string employeeName = ""; 
      DateTime insertingDateTime = new DateTime(1999, 1, 1);
      _noOfDaysInMonth = DateTime.DaysInMonth(_initialDateTime.Year, _initialDateTime.Month);
      while (sourceIndex <= _sourceRowCount)
      {
        employeeName = _ReadExcel.ReadCell(sourceIndex, _nameColumnNo);
        DateTime sourceDateTime =  GetDateTime(_ReadExcel.ReadCell(sourceIndex,_dateColumnNo));
        if (employeeName != _currentEmployeeName || sourceDateTime.Month != _initialDateTime.Month)
        {
          _writeExcel.Save();
          _writeExcel.Close();//Not sure 
          _currentEmployeeName = employeeName;
          _initialDateTime = sourceDateTime;
          CreateNewFile();
          targetIndex = 1;
          insertingDateTime = new DateTime(1999, 1, 1);
          _noOfDaysInMonth = DateTime.DaysInMonth(_initialDateTime.Year, _initialDateTime.Month);
        }
        if (targetIndex > _noOfDaysInMonth)
        {
          targetIndex = _noOfDaysInMonth;
        }
        DateTime targetDateTime =  GetDateTime(_writeExcel.ReadCell(6+targetIndex,2));
        if (sourceDateTime.Date>targetDateTime.Date)
        {
          if (insertingDateTime.Date == targetDateTime.Date)
          {
            _writeExcel.WriteToCell(6+targetIndex,4,"");
            targetIndex++;
          }
          else
          {
            _writeExcel.WriteToCell(6+targetIndex,3,"");
            _writeExcel.WriteToCell(6+targetIndex,4,"");
            targetIndex++;
          }
        }
        else if(sourceDateTime.Date == targetDateTime.Date)
        {
          if (sourceDateTime.Date != insertingDateTime.Date)
          {
            _writeExcel.WriteToCell(6+targetIndex,3,sourceDateTime.ToString("HH:mm"));
            insertingDateTime = sourceDateTime;
            sourceIndex++;
          }
          else if(sourceDateTime.Date == insertingDateTime.Date)
          {
            _writeExcel.WriteToCell(6+targetIndex,4,sourceDateTime.ToString("HH:mm"));
            sourceIndex++;
            targetIndex++;
          }
        }
        else if (sourceDateTime.Date<targetDateTime.Date)
        {
          _writeExcel.WriteToCell(6+(targetIndex-1),4,sourceDateTime.ToString("HH:mm"));
          sourceIndex++;
        }
        Console.WriteLine("RUNNING");
      }
      _writeExcel.Save();
      Console.WriteLine("Multiple file is created for floor 6");
    }
  }
}
