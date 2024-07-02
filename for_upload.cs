Easiest way to read text file which is locked by another application
https://stackoverflow.com/questions/1389155/easiest-way-to-read-text-file-which-is-locked-by-another-application


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
using static System.Net.WebRequestMethods;
using System.Xml;

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
        [ExcelCommand(MenuName = "Test", MenuText = "Yoji Test1", ShortCut = "^Q")]
        public static void YojiTest1(object values)
        {
            dynamic xlApp = ExcelDnaUtil.Application;

            // If you do
            // Excel.Application xlApp2 = (Excel.Application)ExcelDnaUtil.Application;
            // YOU WILL LOSE HOT RELOAD!!!
            //https://stackoverflow.com/questions/75570571/hot-reload-is-not-available-due-to-cs7096
            //Hot relead only works for fully managed code. It patches the IL in the runtime. It can't do that with COM objects

            //Need to remove (error) reference to VBAStreamDecompress.dll, if dll doesn't exist. This prevented using HOT RELOAD.

            //xlCalculationAutomatic -4105 xlCalculationManual - 4135 xlCalculationSemiautomatic 2

            //xlApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual; This will prevent using HOT RELOAD
            xlApp.Calculation = -4135;
            xlApp.EnableEvents = false;
            xlApp.ScreenUpdating = false;

            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            // Get workbook full path
            string sWKBKname = "";
            string wkbkFullPath = "";
            object x = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            object[,] sSheetnames = (object[,])XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            string ssss = sSheetnames[0, 0].ToString();
            string sh_name_only = ssss.Substring(1 + ssss.IndexOf("]"));
            var sPath = XlCall.Excel(XlCall.xlfGetDocument, 2, sh_name_only); // returns Error when workbook is not saved yet
            if (sPath.ToString().IndexOf(":") < 0)
            {
                ;// return; /* ERROR */
            }
            else
            {
                sWKBKname = (string)XlCall.Excel(XlCall.xlfGetDocument, 88, sh_name_only);
                wkbkFullPath = sPath + "\\" + sWKBKname;
            }
            if (false && sWKBKname.Length > 0)
            {
                // Get hash of VBA code
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
                    //var hash = VBAStreamDecompress.MVGvbaDecompress.archiveVBAcodes2(new MemoryStream(fileBytes), "");
                }
                catch (Exception)
                {
                    throw;
                }
                return;
            }

            ExcelReference inRange;
            object result = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            object[,] sheetNames = (object[,])result;

            /////////////////////////////////////////////////////
            ///
            int iii = 69;
            string rangeText2 = xlApp.Worksheets[1].usedrange.address;
            //xlApp.Worksheets[1].cells(1,1)
            //xlApp.Worksheets[1].cells(1,1).locked
            //If there is only one filled cell, return is string otherwise object[,]
            //object[,] formulaArray1 = (object[,])xlApp.Worksheets[1].usedrange.Formula;
            dynamic formulaArray1 = xlApp.Worksheets[1].usedrange.Formula;
            //if(formulaArray1 is System.Object[,])
            if (formulaArray1 is System.String) { } else { }


            //ExcelReference sheetRef = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, "xxxxxxx");
            //testFind(xlApp, (string)sheetNames[0, 0], 1, 1); // This prevents hot reload
            /////////////////////////////////////////////////////

            int distinctFormulaCount = 0;
            var strFormulas = new StringBuilder();
            var sheet_count = xlApp.Worksheets.Count;
            for (int j = 0; j < sheet_count; j++)
            {
                string sheetName = (string)sheetNames[0, j];
                string rangeText = xlApp.Worksheets[j + 1].usedrange.address; //COM?
                double lastRow = (double)XlCall.Excel(XlCall.xlfGetDocument, 10, sheetName); // C API (usedRange) Number of the last used row. If the sheet is empty, returns 0.
                double lastCol = (double)XlCall.Excel(XlCall.xlfGetDocument, 12, sheetName);

                var distinctFormulae = new HashSet<string>();

                for (int i1 = 0; i1 < lastRow; i1++)
                {
                    for (int i2 = 0; i2 < lastCol; i2++)
                    {
                        inRange = new ExcelReference(i1, i1, i2, i2, sheetName);
                        bool flg1 = (bool)XlCall.Excel(XlCall.xlfGetCell, 48, inRange); // Isformula:true/false
                        bool flg2 = (bool)XlCall.Excel(XlCall.xlfGetCell, 14, inRange); // If the cell is locked, returns TRUE; otherwise, returns FALSE.

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

            string hashStr = calcHash(strFormulas.ToString());

            sw.Stop(); TimeSpan ts = sw.Elapsed;
            string lapse = ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00") + ":" + ts.Seconds.ToString("00");
            MessageBox.Show("Time:" + lapse + "\n" + "distinctFormulaCount:" + distinctFormulaCount.ToString() + "\n" + hashStr);


            xlApp.ScreenUpdating = true;
            xlApp.EnableEvents = true;
            xlApp.Calculation = -4105;

        }

        private static string calcHash(string targetStr)
        {
            //var targetStr = "example";
            //targetStr = new string('*', 5000);

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

            return hashStr.ToString();
        }

//Useless function. Too Slow.
        private static void testFind(dynamic xlApp, string sheetName, int lastRow, int lastCol)
        {
            //If parameter is declared as "Excel.Application xlApp" then hot reload will fail!
            int ii = 9;
            ExcelReference inRange = new ExcelReference(0, lastRow - 1, 0, lastCol - 1, sheetName);
            //var str1 = (string)XlCall.Excel(XlCall.xlfReftext, inRange);

            //https://stackoverflow.com/questions/17387443/fast-method-for-determining-unlocked-cell-range

            //ExcelReference FoundCell; // inRange = new ExcelReference(i1, i1, i2, i2, sheetName);

            string FirstCellAddr;
            //ExcelReference UnlockedUnion=null;
            Excel.Range UnlockedUnion = null;

            //'NOTE: When finding by format, you must first set the FindFormat specification:

            xlApp.FindFormat.Clear();
            xlApp.FindFormat.Locked = true;

            //'NOTE: Unfortunately, the FindNext method does not remember the SearchFormat:=True specification so it is
            //'necessary to capture the address of the first cell found, use the Find method (instead) inside the find-next
            //'loop and explicitly terminate the loop when the first-found cell is found a second time.

            //error FoundCell = xlApp.Worksheets[1].usedRange.Find("1", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, true);
            Excel.Range FoundCell = (Excel.Range)xlApp.Worksheets[1].usedRange.Find("1", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, true);
            Excel.Range FirstCell = null;
            //FoundCell = xlApp.Worksheets[1].usedrange.Find(What:= "", After:=.Cells(1, 1), LookIn:= xlFormulas, LookAt:= xlPart,
            //                      SearchOrder:= xlByRows, SearchDirection:= xlNext, MatchCase:= False,
            //                      SearchFormat:= True);

            

            if (FoundCell != null)
            {
                var first_address = FoundCell.Address[true, true, Excel.XlReferenceStyle.xlA1, true, null];
                //FirstCell = FoundCell;
                do
                {
                    //var xxx = FoundCell.GetType().GetProperties();
                    //Debug.Print FoundCell.Address
                    if (UnlockedUnion == null)
                    {
                        UnlockedUnion = FoundCell.MergeArea;//                         'Include merged cells, if any
                    }
                    else
                    {
                        UnlockedUnion = xlApp.Union(UnlockedUnion, FoundCell.MergeArea);
                    }
                    FoundCell = (Excel.Range)xlApp.Worksheets[1].usedRange.Find("1", FoundCell, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, true);
                } while (first_address != FoundCell.Address[true, true, Excel.XlReferenceStyle.xlA1, true, null]);

            }
            xlApp.FindFormat.Clear();

            //Set GetUnlockedCells = UnlockedUnion


        }
    }
}
