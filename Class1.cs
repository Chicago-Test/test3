using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel; //using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Security.Policy;

using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.IO;

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
            return str + "english";
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
        [ExcelCommand(MenuName = "Test", MenuText = "Range Set2", ShortCut = "^Q")]
        public static void RangeSet2(object values)
        {
            dynamic xlApp = ExcelDnaUtil.Application;

            // If you do
            // Excel.Application xlApp2 = (Excel.Application)ExcelDnaUtil.Application;
            // YOU WILL LOSE HOT RELOAD!!!
            //https://stackoverflow.com/questions/75570571/hot-reload-is-not-available-due-to-cs7096
            //Hot relead only works for fully managed code. It patches the IL in the runtime. It can't do that with COM objects

            //xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();


            // Get workbook full path
            object x = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            object[,] sSheetnames = (object[,])XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            string ssss = sSheetnames[0, 0].ToString();
            string sh_name_only = ssss.Substring(1 + ssss.IndexOf("]"));
            var sPath = XlCall.Excel(XlCall.xlfGetDocument, 2, sh_name_only); // returns Error when workbook is not saved yet
            if (sPath.ToString().IndexOf(":") < 0) { return; /* ERROR */ }
            var sWKBKname = XlCall.Excel(XlCall.xlfGetDocument, 88, sh_name_only);
            string wkbkFullPath = sPath + "\\" + sWKBKname;

            byte[] fileBytes;
            try
            {
                //Can read a file used by another process
                using (var fs = new FileStream(wkbkFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    fileBytes = new byte[fs.Length];
                    fs.Read(fileBytes, 0, fileBytes.Length);
                    fs.Close();
                }
                // Need reference to VBAStreamDecompress.dll, and need to copy 7z.dll to "x86" or "x64"
                var hash = VBAStreamDecompress.MVGvbaDecompress.archiveVBAcodes2(new MemoryStream(fileBytes), "");
            }
            catch (Exception)
            {
                throw;
            }

            return;

            ExcelReference inRange;
            object result = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            object[,] sheetNames = (object[,])result;

            /////////////////////////////////////////////////////
            ///
            int iii = 67;
            string rangeText2 = xlApp.Worksheets[1].usedrange.address;
            //xlApp.Worksheets[1].cells(1,1)
            //xlApp.Worksheets[1].cells(1,1).locked
            //If there is only one filled cell, return is string otherwise object[,]
            //object[,] formulaArray1 = (object[,])xlApp.Worksheets[1].usedrange.Formula;
            dynamic formulaArray1 = xlApp.Worksheets[1].usedrange.Formula;
            //if(formulaArray1 is System.Object[,])
            if (formulaArray1 is System.String) { } else { }


            ExcelReference sheetRef = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, "xxxxxxx");

            // Inside testFind, hot reload does not work!!!
            testFind(xlApp, (string)sheetNames[0, 0], 1, 1);
            /////////////////////////////////////////////////////

            int distinctFormulaCount = 0;
            var strFormulas = new StringBuilder();
            var sheet_count = xlApp.Worksheets.Count;
            for (int j = 0; j < sheet_count; j++)
            {
                string sheetName = (string)sheetNames[0, j];
                string rangeText = xlApp.Worksheets[j + 1].usedrange.address; //COM
                double lastRow = (double)XlCall.Excel(XlCall.xlfGetDocument, 10, sheetName); // C API (usedRange) Number of the last used row. If the sheet is empty, returns 0.
                double lastCol = (double)XlCall.Excel(XlCall.xlfGetDocument, 12, sheetName);

                //testFind(xlApp, sheetName, (int)lastRow, (int)lastCol);


                var distinctFormulae = new HashSet<string>();

                for (int i1 = 0; i1 < lastRow; i1++)
                {
                    for (int i2 = 0; i2 < lastCol; i2++)
                    {
                        inRange = new ExcelReference(i1, i1, i2, i2, sheetName);
                        bool flg1 = (bool)XlCall.Excel(XlCall.xlfGetCell, 48, inRange); // Isformula:true/false
                        if (flg1 == true)
                        {
                            //Excel makes it tricky to make a function like that which reads the formula.
                            //Excel requires IsMacroType = true for you to call xlfGetFormula, but then has the side effect of becoming volatile.
                            //dynamic xx2 = (string)XlCall.Excel(XlCall.xlfGetFormula, inRange); // formula in R1C1-style references
                            dynamic xx3 = (string)XlCall.Excel(XlCall.xlfGetCell, 6, inRange); //Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.
                            strFormulas.Append(j.ToString("000") + i1.ToString("0000000") + i2.ToString("00000") + xx3);
                            distinctFormulae.Add(xx3.ToString());
                        }
                    }
                }
                distinctFormulaCount += distinctFormulae.Count();
            }

            XlCall.Excel(XlCall.xlcOptionsGeneral, 1); //Use 1 for A1 style references; 2 for R1C1 style references

            // hash
            var targetStr = "example";
            targetStr = new string('*', 5000);

            targetStr = strFormulas.ToString();
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

            sw.Stop(); TimeSpan ts = sw.Elapsed;
            string lapse = ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00") + ":" + ts.Seconds.ToString("00");
            MessageBox.Show("Time:" + lapse + "\n" + "distinctFormulaCount:" + distinctFormulaCount.ToString() + "\n" + hashStr.ToString());

        }
        private static void testFind(dynamic xlApp, string sheetName, int lastRow, int lastCol)
        {
            //If parameter is declared as "Excel.Application xlApp" then hot reload will fail!
            int ii = 9;
            //ExcelReference inRange = new ExcelReference(0, lastRow - 1, 0, lastCol - 1, sheetName);
            //var str1 = (string)XlCall.Excel(XlCall.xlfReftext, inRange);
        }
    }
}
