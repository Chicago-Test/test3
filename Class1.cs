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
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            //xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            ExcelReference inRange;
            inRange = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            string refText = (string)XlCall.Excel(XlCall.xlfReftext, inRange, true); // returns excel address "[Book1]Sheet1!$E$5:$F$8"
            dynamic range1 = xlApp.Range[refText]; //range1[1].address

            //Range targetRange = xlApp.activeworkbook.Sheets[1].Range["A1:C2"];
            //Range dateCell = targetRange.Cells[2, 2];


            Excel.Workbook wb = xlApp.ActiveWorkbook;
            object[,] dat = new object[3, 4];
            //xlApp.Sheets[1].Range(xlApp.Sheets[1].Cells[1, 1], xlApp.Sheets[1].Cells[1000, 256]).value = dat;
            object[,] formulaArr = xlApp.Sheets[1].range(xlApp.Sheets[1].cells(1, 1), xlApp.Sheets[1].cells(10, 5)).formula;


            //dynamic flg = XlCall.Excel(XlCall.xlcSelectSpecial,3,23,1); // does not work on protected sheet
            inRange = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

            try
            {
                inRange = new ExcelReference(2, 2, 1, 1, "Sheet1");
                dynamic xx = XlCall.Excel(XlCall.xlfGetCell, 48, inRange); // Isformula:true/false
                dynamic xx1 = XlCall.Excel(XlCall.xlfGetFormula, inRange); // formula in R1C1-style references
                xlApp.Range["F1"].Value = "Testing 1... 2... 3... 4";
            }
            catch (Exception ex)
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

            ExcelReference sheetRef = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetNames[0, 0]);
            string str1 = xlApp.ActiveWorkbook.Worksheets[1].codename; // index starts from one

            dynamic r1c1_mode = XlCall.Excel(XlCall.xlfGetWorkspace, 4); //If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.
            XlCall.Excel(XlCall.xlcOptionsGeneral, 2); //Use 1 for A1 style references; 2 for R1C1 style references
            dynamic xxx4 = XlCall.Excel(XlCall.xlfFormulaConvert, "=SUM(D3:D5)+1+Sheet2!C4", true, false, 1);


            int distinctFormulaCount = 0;
            var strFormulas = new StringBuilder();
            var sheet_count = xlApp.Worksheets.Count;
            for (int j = 0; j < sheet_count; j++)
            {
                string sheetName = (string)sheetNames[0, j];
                string rangeText = xlApp.Worksheets[j + 1].usedrange.address; //COM
                double lastRow = (double)XlCall.Excel(XlCall.xlfGetDocument, 10, sheetName); // C API (usedRange) Number of the last used row. If the sheet is empty, returns 0.
                double lastCol = (double)XlCall.Excel(XlCall.xlfGetDocument, 12, sheetName);

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
            string lapse = ts.Hours.ToString("00")+":"+ts.Minutes.ToString("00")+":"+ts.Seconds.ToString("00");
            MessageBox.Show("Time:" + lapse +"\n"+"distinctFormulaCount:" + distinctFormulaCount.ToString() + "\n" + hashStr.ToString());

        }

    }
}
