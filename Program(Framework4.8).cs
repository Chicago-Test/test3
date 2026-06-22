using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using TreeSitter;

namespace TreeSitterTest
{
    // A class to test the TreeSitter methods
    public class Test
    {
        public static void Main(string[] args)
        {
            string path = @"E:\temp\cloc\uncomment\src";
            var outFile = "$temp.txt";
            long checkParameters = 0;
            bool flagOverWrite = false;
            bool flagReadConfig = false;
            string pathConfigFile = "";
            if (args.Length < 1)
            {
                usage("TreeSitterTest.exe");
                Environment.Exit(1);
            }
            for (int i = 0; i < args.Length; i++)
            {
                if (string.Equals(args[i], "-overwrite", StringComparison.OrdinalIgnoreCase)) { flagOverWrite = true; }
                else if (string.Equals(args[i], "-config", StringComparison.OrdinalIgnoreCase))
                {
                    flagReadConfig = true; pathConfigFile = args[i + 1]; pathConfigFile = Path.GetFullPath(pathConfigFile);
                    i++;
                }
                else { path = args[i]; checkParameters |= (1 << 0); }
            }
            if (checkParameters != 1) { usage("TreeSitterTest.exe"); Environment.Exit(1); }
            if (flagOverWrite)
            {
                Console.Write("Are you sure to override the files? Press 'R' to continue the process...");

                // here it ask to press "E" to exit
                if (Console.ReadKey().Key != ConsoleKey.R)
                {
                    Environment.Exit(1);
                }
            }

            var dic1 = new Dictionary<string, string>();
            setExtension(dic1);
            if (flagReadConfig == true)
            {
                try
                {
                    dic1.Clear();
                    var records = File.ReadLines(pathConfigFile);
                    foreach (var record in records)
                    {
                        if (record.Length > 0)
                        {
                            if (record.TrimEnd().Substring(0, 1).CompareTo("#") == 0) { continue; }
                            var values = record.Split(',');
                            dic1.Add("." + values[0], values[1]);
                            //Console.WriteLine(record);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    Environment.Exit(1);
                }
            }

            //foreach (var key in dic1.Keys)
            //{
            //    Console.WriteLine($"{key} : {dic1[key]}");
            //}

            /////////////////////////////
            //string[] files = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories);
            var files = Directory.EnumerateFiles(path, "*.*", SearchOption.AllDirectories); //.ToArray();
                                                                                            //string exactPath = Path.GetFullPath(path);
            StringBuilder sb = new StringBuilder();
            int count = 0;
            foreach (string file in files)
            {
                //Console.WriteLine(file);
                string extension = Path.GetExtension(file);
                ///////////////////////////////////////////////////////////////////////////////////////////
                if (dic1.ContainsKey(extension))
                {
                    count++;
                    Console.WriteLine(file);
                    try
                    {
                        var lang = dic1[extension]; //c-sharp, javascript, vb_dotnet DLL name must match:"tree-sitter-<lang>.dll"
                                                    //var fileFullPath = "E:\\temp\\cloc\\AfLMM1_utf8.cs";
                                                    //var fileFullPath = "C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319\\SQL\\ja\\SqlWorkflowInstanceStoreSchemaUpgrade.sql";
                        var fileFullPath = file;
                        using (var language = new Language(lang))
                        {
                            using (var parser = new Parser(language))
                            {
                                var filetext = File.ReadAllText(fileFullPath);
                                //using var tree = parser.Parse("function one() { function two() {} }")!;
                                using (var tree = parser.Parse(filetext))
                                {
                                    if (tree != null)
                                    {
                                        //if (variable == null)

                                        //var rangesToRemove = new List<(int Start, int End)>();
                                        var rangesToRemove = new List<RngToRemove>();

                                        using (var query = new Query(language, "(comment) @comment"))
                                        {
                                            foreach (var capture in query.Execute(tree.RootNode).Captures)
                                            {
                                                //Console.WriteLine($"Found function: {capture.Node.Text}");
                                                //rangesToRemove.Add((capture.Node.StartIndex, capture.Node.EndIndex)); //capture.Node.EndPosition
                                                rangesToRemove.Add(new RngToRemove(capture.Node.StartIndex, capture.Node.EndIndex,-capture.Node.StartIndex));
                                            }
                                        }
                                        // 5. Delete text from back to front to avoid altering subsequent indices
                                        sb = new StringBuilder(filetext);
                                        //foreach (var range in rangesToRemove.OrderByDescending(r => r.Start))
                                        foreach (var range in rangesToRemove.OrderBy(r => r.SignReversedStart))
                                        {
                                            int length = range.End - range.Start;
                                            sb.Remove(range.Start, length);
                                        }
                                    }
                                }
                            }
                        }

                        var str = sb.ToString();
                        if (flagOverWrite == true)
                        {
                            using (StreamWriter sw = File.CreateText(fileFullPath))
                            {
                                sw.Write(str);
                                sw.Close();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        Environment.Exit(1);
                    }
                }
            }
            if (flagOverWrite)
            {
                Console.WriteLine("Total " + count + " files were overwritten.");
            }
            else
            {
                Console.WriteLine("Total " + count + " files will be overwritten.");
            }
        }

        private static void setExtension(Dictionary<string, string> dic1)
        {
            dic1.Add(".bas", "vb_dotnet");
            dic1.Add(".frm", "vb_dotnet");
            dic1.Add(".vb", "vb_dotnet");
            dic1.Add(".r", "r");
            dic1.Add(".sql", "sql");
            dic1.Add(".c", "h");
            dic1.Add(".h", "c");
        }
        private static void usage(string prg)
        {
            Console.WriteLine("Remove comment lines (including in-line comments) from source codes.");
            Console.WriteLine("List all targetted files: >{0} \"C:\\models\\src\"", prg);
            Console.WriteLine("Remove comment lines: >{0} \"C:\\models\\src\" -overwrite", prg);
            Console.WriteLine("Remove comment lines: >{0} \"C:\\models\\src\" -config <config file name>", prg);
            Console.WriteLine("Config file is comma separated csv. <Extension>,<Partial DLL name>");
            Console.WriteLine("    (E.g. tree-sitter-vb_dotnet.dll -> vb_dotnet)");
            Console.WriteLine("    bas,vb_dotnet");
            Console.WriteLine("    js,javascript");

        }
        private class RngToRemove
        {
            public int Start { get; set; }
            public int End { get; set; }
            public int SignReversedStart { get; set; }
            public RngToRemove(int start,int end, int signReversedStart)
            {
                Start = start;
                End = end;
                SignReversedStart = signReversedStart; // For OrderByDescending
            }
        }
    }
}
