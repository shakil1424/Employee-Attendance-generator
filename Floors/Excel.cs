using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace Floors
{
  public class Excel
  {
    public string path { get; set; }
    _Application excel = new _Excel.Application();
    private Workbook wb;
    private Worksheet ws;


    public Excel(string path, int sheet)
    {
      this.path = path;
      wb = excel.Workbooks.Open(path);
      ws = wb.Worksheets[sheet];
    }

    public Excel()
    {

    }

    public string ReadCell(int i, int j)
    {
      /*i++;
      j++;*/
      if (ws.Cells[i, j].Value2 != null)
      {
        string data = ws.Cells[i, j].Value2.ToString();
        return data;
      }

      return "";


    }
    public void WriteToCell(int i, int j, string s)
    {
      ws.Cells[i, j].Value2 = s;
    }

    public int CountUsedRows()
    {
      Range excelCell = ws.UsedRange;
      Object[,] sheetValues = (Object[,])excelCell.Value;
      int count = 0;
      count = sheetValues.GetLength(0);
      return count;
    }
    public int CountUsedColumns()
    {
      Range excelCell = ws.UsedRange;
      Object[,] sheetValues = (Object[,])excelCell.Value;
      int count = 0;
      count = sheetValues.GetLength(1);
      return count;
    }
    public void Save()
    {
      wb.Save();
    }
    public void SaveAs(string path)
    {
      wb.SaveAs(path);
    }
    public void Close()
    {
      wb.Close(0);
    }
  }
}
