using System;
using System.Collections.Generic;
using Utils.OpenXmlService;

namespace Test
{
    public class Program
    {
        public static void Main(string[] args)
        {
            TestExcelRead();
            Console.WriteLine("press any key to continue");
            Console.WriteLine();
            Console.ReadKey();

            if (args.Length > 1)
            {
                var path = args[0];
                var sheetName = args[1];

                TestExcelOpen(path, sheetName);
                Console.WriteLine("press any key to continue");
                Console.WriteLine();
                Console.ReadKey();
            }

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

            /// <summary>
            /// Input array:    a1      b1          c1
            ///                 1       2           3
            ///                 true    1-1-2012    true    true
            /// </summary>
            private static void TestExcelRead()
        {
            var document = new ExcelDocument();
            string[][] array = { new[] { "a1", "b1", "c1" }, new[] { "1", "2", "3" }, new[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            var vs = document.GetRange("Sheet1", "B1", "C3");
            foreach (var strings in vs)
            {
                foreach (var s in strings)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Opens the given sheet of the given workbook from B1 to C10
        /// </summary>
        private static void TestExcelOpen(string path, string sheetName)
        {
            var document = ExcelDocument.Open(path);
            var vs = document.GetRange(sheetName, "B1", "C10");
            foreach (var strings in vs)
            {
                foreach (var s in strings)
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
        private static void TestExcelPaste()
        {
            var document = new ExcelDocument();
            string[][] array = { new[] { "a1", "b1", "c1" }, new[] { "1", "2", "3" }, new[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            var vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (var strings in vs)
            {
                foreach (var s in strings)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();

            }
            Console.WriteLine();
            var a = new List<List<string>>() { new List<string> { "a2", "b2" }, new List<string> { "4", "5" } };
            document.PasteRange("Sheet1", "C3", a);
            vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (var strings in vs)
            {
                foreach (var s in strings)
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
        private static void TestExcelCut()
        {
            var document = new ExcelDocument();
            string[][] array = { new[] { "a1", "b1", "c1" }, new[] { "1", "2", "3" }, new[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            var vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (var strings in vs)
            {
                foreach (var s in strings)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();

            }

            Console.WriteLine();
            var a = document.CutRange("Sheet1", "A1", "B2");
            document.PasteRange("Sheet1", "C3", a);
            vs = document.GetRange("Sheet1", "A1","D4");
            foreach (var strings in vs)
            {
                foreach (var s in strings)
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
        private static void TestExcelScanToEnd()
        {
            var document = new ExcelDocument();
            string[][] array = { new[] { "a1", "b1", "c1" }, new[] { "1", "2", "3" }, new[] { "true", "1-1-2012", "true", "true" } };
            document.AddArray(array, "Sheet1");
            var vs = document.GetRange("Sheet1", "A1", "D4");
            foreach (var strings in vs)
            {
                foreach (var s in strings)
                {
                    Console.Write($"{s}, ");
                }
                Console.WriteLine();
            }

            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.DOWN));//A3
            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.UP, "B3"));//B1
            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.RIGHT));//C1
            Console.WriteLine(document.ScanToEnd("Sheet1", ExcelDocument.Direction.LEFT, "D3"));//A3
        }

        /// <summary>
        /// Output : A7, A1, D4, A4
        /// </summary>
        private static void TestExcelScan()
        {
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.DOWN, "A4", 3));//A7
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.UP, "A4", 5));//A1
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.RIGHT, "A4", 3));//D4
            Console.WriteLine(ExcelDocument.Scan(ExcelDocument.Direction.LEFT, "A4", 3));//A4

        }
    }
}
