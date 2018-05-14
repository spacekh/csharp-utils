using System;
using System.Collections.Generic;
using Utils.OpenXmlService;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            TestExcelRead();
            Console.WriteLine("press any key to continue");
            Console.WriteLine();
            Console.ReadKey();

            //string path, sheetName;
            //if (args.Length > 1) { path = args[0]; sheetName = args[1]; }

            //TestExcelOpen(path, sheetName);
            //Console.WriteLine("press any key to continue");
            //Console.WriteLine();
            //Console.ReadKey();

            TestExcelPaste();
            Console.WriteLine("press any key to continue");
            Console.WriteLine();
            Console.ReadKey();

            TestExcelCut();
            Console.WriteLine("press any key to continue");
            Console.WriteLine();
            Console.ReadKey();

            TestExcelScanToEnd();
            Console.WriteLine("press any key to continue");
            Console.WriteLine();
            Console.ReadKey();

            TestExcelScan();
            Console.WriteLine("press any key to continue");
            Console.WriteLine();
            Console.ReadKey();
        }

        static void TestExcelRead()
        {
            /// <summary>
            /// Input array:    a1      b1          c1
            ///                 1       2           3
            ///                 true    1-1-2012    true    true
            /// </summary>
            ExcelDocument document = new ExcelDocument();
            string[][] array = { new string[] { "a1", "b1", "c1" }, new string[] { "1", "2", "3" }, new string[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            List<List<string>> vs = document.GetRange("Sheet1", "B1", "C3");
            foreach (List<string> strs in vs)
            {
                foreach (string s in strs)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Opens the given sheet of the given workbook from B1 to C10
        /// </summary>
        static void TestExcelOpen(string path, string sheetName)
        {
            ExcelDocument document = ExcelDocument.Open(path);
            List<List<string>> vs = document.GetRange(sheetName, "B1", "C10");
            foreach (List<string> strs in vs)
            {
                foreach (string s in strs)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Input array:    a1      b1          c1                  Output: a1      b1                  c1
        ///                 1       2           3                           1       2                   3
        ///                 true    1-1-2012    true    true                TRUE    January 1, 2012     a2  b2
        ///                                                                 []      []                  4   5
        /// </summary>
        static void TestExcelPaste()
        {
            ExcelDocument document = new ExcelDocument();
            string[][] array = { new string[] { "a1", "b1", "c1" }, new string[] { "1", "2", "3" }, new string[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            List<List<string>> vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (List<string> strs in vs)
            {
                foreach (string s in strs)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();

            }
            Console.WriteLine();
            List<List<string>> a = new List<List<string>>() { new List<string> { "a2", "b2" }, new List<string> { "4", "5" } };
            document.PasteRange("Sheet1", "C3", a);
            vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (List<string> strs in vs)
            {
                foreach (string s in strs)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();

            }
        }


        /// <summary>
        /// Input array:    a1      b1          c1                  Output: []      []                  c1
        ///                 1       2           3                           []      []                  3
        ///                 true    1-1-2012    true    true                TRUE    January 1, 2012     a1  b1
        ///                                                                 []      []                  1   2
        /// </summary>
        static void TestExcelCut()
        {
            ExcelDocument document = new ExcelDocument();
            string[][] array = { new string[] { "a1", "b1", "c1" }, new string[] { "1", "2", "3" }, new string[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            List<List<string>> vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (List<string> strs in vs)
            {
                foreach (string s in strs)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();

            }

            Console.WriteLine();
            List<List<string>> a = document.CutRange("Sheet1", "A1", "B2");
            document.PasteRange("Sheet1", "C3", a);
            vs = document.GetRange("Sheet1", "A1","D4");
            foreach (List<string> strs in vs)
            {
                foreach (string s in strs)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();

            }
        }

        /// <summary>
        /// Input array:    a1      b1          c1                  Output: A3
        ///                 1       2           3                           B1
        ///                 true    1-1-2012    true    true                C1
        ///                                                                 A3
        ///</summary>
        static void TestExcelScanToEnd()
        {
            ExcelDocument document = new ExcelDocument();
            string[][] array = { new string[] { "a1", "b1", "c1" }, new string[] { "1", "2", "3" }, new string[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            List<List<string>> vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (List<string> strs in vs)
            {
                foreach (string s in strs)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();
            }

            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.down));//A3
            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.up, "B3"));//B1
            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.right));//C1
            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.left, "D3"));//A3
        }

        /// <summary>
        /// Output : A7, A1, D4, A4
        /// </summary>
        static void TestExcelScan()
        {
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.down, "A4", 3));//A7
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.up, "A4", 5));//A1
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.right, "A4", 3));//D4
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.left, "A4", 3));//A4

        }
    }
}
