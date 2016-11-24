using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
namespace HOffice
{
    public class Excel
    {
        Application app = new Application();
        Workbook book;
        Worksheet sheet;
        string path;
        public Excel(string path)
        {
            this.path = path;
        }
        public bool Open()
        {
            book = app.Workbooks.Open(path);
            this.sheet = (Worksheet)book.Sheets[1];
            return true;
        }
        public bool Quit()
        {
            app.Quit();
            return true;
        }
        public bool Write(string value, int column = 1,int row = 1, int sheet = 1)
        {
            this.sheet = (Worksheet)book.Sheets[sheet];
            this.sheet.Cells[row, column].Value = value;
            book.Save();
            return true;
        }
        public int lastRow()
        {
            return this.sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
        }
    }
}
