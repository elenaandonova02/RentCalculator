using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Rentu_Kalkulator_Aktuarstvo_2021
{
    class Excel
    {
        string path = "";
        int i = 0;
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public string ReadCell(int i, int j)
        {
            //i++;
            //j++;
            //if (ws.Cells[i, j].Value2 != null)
              //  return ws.Cells[i, j].Value2;
            //else
                return "";
        }

        // do tuka e od videoto, nadole ovoj kod e nov dopisan, ne rabote so toa gore
        //samo vo funkcijata ReadCell ostavi go return "" ; oti error ke ima
        public string getCellValue(int i, int j)
        {
            string cellValue = ws.Cells[i, j].Text.ToString();
            return cellValue;
        }

    }
}
