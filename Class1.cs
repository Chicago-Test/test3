using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Cryptography;

namespace ClassLibrary2
{

    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hi " + name;
        }
        //[ExcelFunction(Description = "My first .NET function2")]
        public static string SayHello2(string name)
        {
            return "Hello " + name;
        }
        [ExcelFunction(IsMacroType = true)]
        public static object[,] ReadFormulasMacroType(
        [ExcelArgument(AllowReference = true)] object arg)
        {
            ExcelReference theRef = (ExcelReference)arg;
            int rows = theRef.RowLast - theRef.RowFirst + 1;
            object[,] res = new object[rows, 1];
            for (int i = 0; i < rows; i++)
            {
                ExcelReference cellRef = new ExcelReference(
                    theRef.RowFirst + i, theRef.RowFirst + i,
                    theRef.ColumnFirst, theRef.ColumnFirst,
                    theRef.SheetId);
                res[i, 0] = XlCall.Excel(XlCall.xlfGetFormula, cellRef);
                
            }
            return res;
        }
        [ExcelFunction(Description = "My first .NET function")]
        public static string ToEnglish([ExcelArgument(Description = "を英訳します。", Name = "対象文字列")] string str)
        {
            return str+"english";
        }
        //Create menu in the Ribbon
        [ExcelCommand(MenuName = "Test", MenuText = "Range Set")]
        public static void RangeSet()
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            
            xlApp.Range["F1"].Value = "Testing 1... 2... 3... 4";

            int i = 1;
            object result = XlCall.Excel(XlCall.xlfGetWorkbook, i);
            object[,] sheetNames = (object[,])result;
            for (int j = 0; j < sheetNames.GetLength(1); j++)
            {
                string sheetName = (string)sheetNames[0, j];
                // use sheetName here.
            }


        }
        //ctrl+shift+Q
        [ExcelCommand(MenuName = "Test",MenuText = "Range Set2", ShortCut = "^Q")]
        public static void RangeSet2(object values)
        {
            dynamic xlApp = ExcelDnaUtil.Application;

            ExcelReference inRange;
            inRange = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            string refText = (string)XlCall.Excel(XlCall.xlfReftext,inRange, true); // returns excel address "[Book1]Sheet1!$E$5:$F$8"
            dynamic range1 = xlApp.Range[refText]; //range1[1].address

            dynamic wb = xlApp.ActiveWorkbook;

            //dynamic flg = XlCall.Excel(XlCall.xlcSelectSpecial,3,23,1); // does not work on protected sheet
            inRange = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

            
            inRange = new ExcelReference(2, 2, 1, 1, "Sheet1");
            dynamic xx = XlCall.Excel(XlCall.xlfGetCell,48,inRange); // Isformula:true/false
            dynamic xx1=XlCall.Excel(XlCall.xlfGetFormula,inRange); // formula in R1C1-style references



            try
            {
                xlApp.Range["F1"].Value = "Testing 1... 2... 3... 4";
            }
            catch (Exception)
            {
        //        throw;
            }

            
            //Application xlapp1=new Application();
            //string str=xlapp1.ActiveSheet.Name;

            //ExcelReference theRef = (ExcelReference)values;

            //MessageBox.Show("ワークシートを新規作成します2");
            //XlCall.Excel(XlCall.xlcWorkbookInsert);
            //XlCall.Excel(XlCall.xlcWorkbookActivate, "Sheet2");
            //https://thinkami.hatenablog.com/entry/20131127/1385503166
            var cell = new ExcelReference(0, 0).GetValue();
            object result = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            object[,] sheetNames = (object[,])result;

            ExcelReference sheetRef =(ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetNames[0,0]);
            string str1=xlApp.activeworkbook.worksheets(1).codename; // index starts from one

            var strFormulas = new StringBuilder();
            var sheet_count = xlApp.Worksheets.Count;
            for (int j = 0; j < sheet_count; j++)
            {
                string sheetName = (string)sheetNames[0, j];
                string rangeText=xlApp.Worksheets[j+1].usedrange.address; //COM
                double lastRow = (double)XlCall.Excel(XlCall.xlfGetDocument, 10,sheetName); // C API (usedRange) Number of the last used row. If the sheet is empty, returns 0.
                double lastCol = (double)XlCall.Excel(XlCall.xlfGetDocument, 12,sheetName);

                for (int i1 = 0; i1 < lastRow; i1++)
                {
                    for (int i2 = 0; i2 < lastCol; i2++)
                    {
                        inRange = new ExcelReference(i1, i1, i2, i2, sheetName);
                        bool flg1 = (bool)XlCall.Excel(XlCall.xlfGetCell, 48, inRange); // Isformula:true/false
                        if (flg1 == true)
                        {
                            dynamic xx2 = (string)XlCall.Excel(XlCall.xlfGetFormula, inRange); // formula in R1C1-style references
                            strFormulas.Append(j.ToString("000")+ i1.ToString("0000000")+ i2.ToString("0000000")+xx2);
                        }
                    }
                }
            }
            // hash
            var targetStr = "example";
            targetStr = new string('*', 5000);
            
            targetStr=strFormulas.ToString();
            var targetBytes = Encoding.UTF8.GetBytes(targetStr);
            // MD5ハッシュを計算
            var csp = new SHA256CryptoServiceProvider();
            var hashBytes = csp.ComputeHash(targetBytes);

            // バイト配列を文字列に変換
            var hashStr = new StringBuilder();
            foreach (var hashByte in hashBytes)
            {
                hashStr.Append(hashByte.ToString("x2"));
            }
            //MessageBox.Show(hashStr.ToString());

        }

    }

}
