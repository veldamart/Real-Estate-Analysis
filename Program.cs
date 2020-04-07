using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Real_Estate_Analysis
{
    class Program
    {
        static void openfile()
        {
          /*this will open and excel file*/
            string path ="C:/Users/Velda/Desktop/Real Estate Analysis/Data/Mortgage payment calculator.xlsx";
            var excelApp = new Application();
            excelApp.Visible = true;

            Excel.Workbooks books = excelApp.Workbooks;

            Excel.Workbook sheet = books.Open(path);
        }
    

        static void Main(string[] args)
        {
            openfile();
        }
    }
}
