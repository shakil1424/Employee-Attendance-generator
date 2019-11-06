using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Floors
{
  public class FifthFloor: FloorNo
  {
    
    public override void CreateMultipleFiles()
    {
      _startingRowIndex = 2;
      _nameColumnNo = 3;
      _dateColumnNo = 5;
      InitializeExcelParameters();
      CreateNewFile();
      int targetIndex = 1;
      int sourceIndex = _startingRowIndex;
      string employeeName = "";
      while (sourceIndex<=_sourceRowCount)
      {
        employeeName = _ReadExcel.ReadCell(sourceIndex, _nameColumnNo);
        DateTime sourceDateTime = GetDateTime(_ReadExcel.ReadCell(sourceIndex, _dateColumnNo));
        if (employeeName != _currentEmployeeName || sourceDateTime.Month != _initialDateTime.Month)
        {
          _writeExcel.Save();
          _writeExcel.Close();
          _currentEmployeeName = employeeName;
          _initialDateTime = sourceDateTime;
          CreateNewFile();
          targetIndex = 1;
        }
        DateTime targetDateTime =  GetDateTime(_writeExcel.ReadCell(6+targetIndex,2));
        if (sourceDateTime.Date > targetDateTime.Date)
        {
          _writeExcel.WriteToCell(6+targetIndex,3,"");
          _writeExcel.WriteToCell(6+targetIndex,4,"");
          targetIndex++;
        }
        else if (sourceDateTime.Date == targetDateTime.Date)
        {
          int timeColumn = 7;
          string dateTime = _ReadExcel.ReadCell(sourceIndex, timeColumn);
          while (dateTime != "")
          {
            DateTime sourceTime = GetDateTime(dateTime);
            if (timeColumn == 7)
            {
              _writeExcel.WriteToCell(6+targetIndex,3,sourceTime.ToString("HH:mm"));
            }
            else
            {
              _writeExcel.WriteToCell(6+targetIndex,4,sourceTime.ToString("HH:mm"));
            }
            timeColumn++;
            dateTime = _ReadExcel.ReadCell(sourceIndex, timeColumn);
          }
          sourceIndex++;
          targetIndex++;
        }
        Console.WriteLine("RUNNING");
      }
      _writeExcel.Save();
      Console.WriteLine("Multiple file is created For Fifth Floor"); 
    }

  }
}
