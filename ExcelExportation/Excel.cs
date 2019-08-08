using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExportation
{
    
    class Excel
    {
        string path = "";
        public _Application excel = new _Excel.Application();
        public Workbook wb;
        public Worksheet ws;

        public Excel()
        {

        }

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
           
        }



        public void CreateNewSheet()
        {
            Worksheet tempSheet = wb.Worksheets.Add(After: ws);

        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 == null)
            {
                return ws.Cells[i, j].Value2;
            }
            else
            {
                return "";
            }
        }

        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnString = new string[endi  - starti, endy-starty];
            for (int p = 1; p <= endi - starti; p++)
            {
                for(int q = 1; q<= endy - starty; q++)
                {
                    if (holder[p, q] == null)
                    {
                        returnString[p - 1, q - 1] = "";
                    }
                    else
                    {
                        returnString[p - 1, q - 1] = holder[p, q].ToString();
                    }
                }
            }
            return returnString;
        }

        public void WriteRange(int starti, int starty, int endi, int endy, string[,]writeString)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writeString;
        }

        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void SelectWorksheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }

        public void DeleteWorkSheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delete();
        }

        public void Close()
        {
            wb.Close();
        }
    }
}