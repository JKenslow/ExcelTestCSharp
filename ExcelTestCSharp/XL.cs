using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestCSharp
{
    public class XL
    {
        private Excel.Application xlApp = new Excel.Application();
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorksheet;
        private Excel.Range xlRange;
        public XL(string fileName)
        {
            xlWorkbook = xlApp.Workbooks.Open(fileName);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
        }

        public string GetCell(int i, int j)
        {
            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            {
                return xlRange.Cells[i, j].Value2.ToString();
            }
            return "";
        }

        public void Save()
        {

        }

        public void SaveAs(string FileName)
        {

        }
    }
}
