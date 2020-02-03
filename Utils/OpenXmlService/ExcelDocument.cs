using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.IO;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Linq;
using System.Text.RegularExpressions;

namespace Utils.OpenXmlService
{
    /// <summary>
    /// A wrapper for <see cref="SpreadsheetDocument"/> that provides multiple methods for easy creation, read/write, and interaction with the document
    /// </summary>
    public class ExcelDocument
    {
        #region Fields
        private SpreadsheetDocument _doc;
        private WorkbookPart _workbookPart => _doc.WorkbookPart;
        private Sheets _sheets => _workbookPart.Workbook.GetFirstChild<Sheets>();
        private readonly MemoryStream _stream = new MemoryStream();
        private static uint _numSheets;
        #endregion

        #region DataTypes
        private static readonly EnumValue<CellValues> NUM_TYPE = new EnumValue<CellValues>(CellValues.Number);
        private static readonly EnumValue<CellValues> STR_TYPE = new EnumValue<CellValues>(CellValues.String);
        private static readonly EnumValue<CellValues> BOOL_TYPE = new EnumValue<CellValues>(CellValues.Boolean);
        private static readonly EnumValue<CellValues> DATE_TYPE = new EnumValue<CellValues>(CellValues.Date);

        #endregion

        #region Initialization
        /// <summary>
        /// Initializes a new, empty Excel document for exporting. (Writes to a MemoryStream instead of a file)
        /// </summary>
        public ExcelDocument()
        {
            //have to create a doc to hold everything, then finish init        
            _doc = SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook);
            Init();
        }

        /// <summary>
        /// Initializes a new, empty Excel document to be written to the given file path.
        /// </summary>
        /// <param name="file">The (full or relative) file path to write the document to</param>
        public ExcelDocument(string file)
        {
            //create a doc using the given filename, then finish init
            _doc = SpreadsheetDocument.Create(file, SpreadsheetDocumentType.Workbook);
            Init();
        }

        /// <summary>
        /// Opens the given file in read/write mode
        /// </summary>
        /// <param name="file">The full filepath for the file to be opened</param>
        /// <returns>A fully initialized ExcelDocument object connected to the given file</returns>
        public static ExcelDocument Open(string file)
        {
            ExcelDocument result = new ExcelDocument
            {
                _doc = SpreadsheetDocument.Open(file, true)
            };
            return result;
        }

        /// <summary>
        /// Common initialization for ExcelDocument class
        /// </summary>
        private void Init()
        {
            //doc holds a WorkbookPart which holds "Sheets" and a WorksheetPart, which in turn holds a Worksheet, which holds a Sheet
            //see ./doc for a diagram of a spreadsheet document
            _doc.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();
            _workbookPart.Workbook.AppendChild(new Sheets());
            _doc.AddExtendedFilePropertiesPart();
        }
        #endregion

        #region Data Import
        /// <summary>
        /// Adds the information from the given DataTable into a new Sheet in the workbook
        /// </summary>
        /// <param name="data">The DataTable to serialize</param>
        public void AddDataTable(DataTable data)
        {
            AddDataTable(data, null);
        }

        /// <summary>
        /// Adds the information from the given DataTable into a new Sheet in the workbook, using the given column ids and alternate names as headers
        /// </summary>
        /// <param name="data">The DataTable to serialize</param>
        /// <param name="cols">The column headers to use.  If not null, only the specified columns, in the specified order, will be written.</param>
        public void AddDataTable(DataTable data, string[,] cols)
        {
            string sheetName = "Sheet" + _numSheets;
            AddDataTable(data, cols, sheetName);
        }

        /// <summary>
        /// Adds the information from the given DataTable into a new Sheet in the workbook, using the given column ids and alternate names as headers, and the given sheet name as the new Sheet's name
        /// </summary>
        /// <param name="data">The DataTable to serialize</param>
        /// <param name="cols">The column headers to use.  If not null, only the specified columns, in the specified order, will be written.</param>
        /// <param name="sheetName">The name to give to the Sheet</param>
        public void AddDataTable(DataTable data, string[,] cols, string sheetName)
        {
            WorksheetPart wspart = InsertSheet(sheetName);
            int col = 0, row = 0;
            string[] head = null;
            //write headers
            try
            {
                //only write the specified headers
                int l = cols.GetLength(0); //throws an exception if cols is null, so if(cols==null) is unnecessary
                head = new string[l];
                for (col = 0; col < l; col++)
                {
                    head[col] = cols[col, 0];
                    Insert(col, row, cols[col, 1], wspart);
                }
            }
            catch (Exception)
            {
                head = new string[data.Columns.Count];
                //write all headers
                foreach (DataColumn c in data.Columns)
                {
                    Insert(col, row, c.ToString(), wspart);
                    head[col] = c.ToString();
                    col++;
                }
            }
            row++;
            //write data
            object o;
            foreach (DataRow r in data.Rows)
            {
                col = 0;

                foreach (string str in head)
                {
                    o = r[head[col]];
                    Insert(col, row, o, wspart);
                    col++;
                }
                row++;
            }

            _workbookPart.Workbook.Save();
        }

        /// <summary>
        /// Adds the information from the given DataTable into a new Sheet in the workbook, using the given column ids and alternate names as headers, and the given sheet name as the new Sheet's name
        /// </summary>
        /// <param name="data">The String array to serialize</param>
        /// <param name="sheetName">The name to give to the Sheet</param>
        public void AddArray(string[][] data, string sheetName)
        {
            WorksheetPart wspart = InsertSheet(sheetName);
            int col = 0, row = 0;

            //write data
            foreach (string[] r in data)
            {
                col = 0;

                foreach (object o in r)
                {
                    Insert(col, row, o, wspart);
                    col++;
                }
                row++;
            }

            _workbookPart.Workbook.Save();
        }

        private WorksheetPart InsertSheet(string sheetName)
        {
            WorksheetPart wspart = _workbookPart.AddNewPart<WorksheetPart>();
            wspart.Worksheet = new Worksheet(new SheetData());

            Sheet sheet = new Sheet();
            _numSheets++;
            sheet.Name = sheetName;
            sheet.Id = _doc.WorkbookPart.GetIdOfPart(wspart);
            sheet.SheetId = _numSheets;
            _sheets.Append(sheet);
            return wspart;
        }

        /// <summary>
        /// Adds all of the DataTables in the given DataSet as separate Sheets in the workbook
        /// </summary>
        /// <param name="ds">The DataSet to serialize</param>
        public void AddDataSet(DataSet ds)
        {
            foreach (DataTable dt in ds.Tables)
            {
                AddDataTable(dt);
            }
        }

        /// <summary>
        /// Inserts the given object into the given (0-indexed) column of the given Row
        /// </summary>
        /// <param name="col">0-indexed column to add this value into</param>
        /// <param name="row">the 0-indexed row to add a cell into</param>
        /// <param name="o">the object to add to the given cell</param>
        /// <param name="wspart">the Worksheet to insert into</param>
        private void Insert(int col, int row, object o, WorksheetPart wspart)
        {
            uint r = (uint)row + 1;
            Cell cell = InsertCellInWorksheet(ColumnNumToName(col), r, wspart);
            Insert(ColumnNumToName(col) + (uint)row + 1, o, wspart);
        }
        
        /// <summary>
        /// Inserts the given object into the given <see cref="CellRef"/> in the given <see cref="WorksheetPart"/>
        /// </summary>
        /// <param name="cr"></param>
        /// <param name="o"></param>
        /// <param name="wspart"></param>
        private static void Insert(CellRef cr, object o, WorksheetPart wspart)
        {
            Cell cell = InsertCellInWorksheet(cr.Col, cr.Row, wspart);
            if (int.TryParse(o.ToString(), out int x) || double.TryParse(o.ToString(), out double y)) cell.DataType = NUM_TYPE;
            else if (bool.TryParse(o.ToString(), out bool z)) cell.DataType = BOOL_TYPE;
            else if (DateTime.TryParse(o.ToString(), out DateTime d)) cell.DataType = DATE_TYPE;
            else cell.DataType = STR_TYPE;

            cell.CellValue = new CellValue(o.ToString());
        }
        
        /// <summary>
        /// Given a column name and a row index, inserts a cell into the worksheet. 
        /// </summary>
        /// <param name="columnName">The column to insert into</param>
        /// <param name="rowIndex">The row to insert into</param>
        /// <param name="wspart">The Worksheet to insert into</param>
        /// <returns>Either the already-existing Cell at the given index, or a new Cell that has been inserted into the worksheet</returns>
        /// <remarks>Credits: https://docs.microsoft.com/en-us/office/open-xml/how-to-insert-text-into-a-cell-in-a-spreadsheet#code-snippet-9 </remarks>
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart wspart)
        {
            SheetData sheetData = wspart.Worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row = null;
            foreach (Row r in sheetData.Elements<Row>())
            {
                if (r.RowIndex == rowIndex)
                {
                    row = r;
                    break;
                }
            }
            if (row == null)
            {
                row = new Row
                {
                    RowIndex = rowIndex
                };
                sheetData.Append(row);
            }

            // If there is already a cell with the specified column name, return it.
            foreach (Cell c in row.Elements<Cell>())
            {
                if (c.CellReference.Value == cellReference) return c;
            }
            //Otherwise, create and insert a new one.
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            int refCellCol = -1, newCellCol = ColumnNameToNum(cellReference);
            foreach (Cell cell in row.Elements<Cell>())
            {
                refCellCol = ColumnNameToNum(cell.CellReference.Value);
                refCell = cell;
                if (refCellCol > newCellCol) break;
            }
            Cell newCell = new Cell();
            newCell.CellReference = cellReference;
            if (refCellCol > newCellCol)
            {
                row.InsertBefore(newCell, refCell);
            }
            else row.InsertAfter(newCell, refCell);
            return newCell;
        }
        #endregion

        #region Read / Write Functions

        /// <summary>
        /// Returns a list of lists of strings corresponding to the cells that were found in the given range. If a worksheet called <paramref name="worksheetName"/> is not found, the first worksheet in the document is searched.
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public List<List<string>> GetRange(string worksheetName, CellRef start, CellRef end)
        {
            List<List<string>> result = new List<List<string>>();
            IEnumerable<Sheet> sheets = _workbookPart.Workbook.Descendants<Sheet>();
            WorksheetPart wspart;
            try
            {
                string relId = sheets.First(s => worksheetName.Equals(s.Name)).Id;
                wspart = (WorksheetPart)_workbookPart.GetPartById(relId);
            }
            catch (InvalidOperationException) { wspart = _workbookPart.GetPartsOfType<WorksheetPart>().First(); }
            SheetData sheetData = wspart.Worksheet.GetFirstChild<SheetData>();

            List<string> values;
            for (uint i = start.Row; i <= end.Row; i++)
            {
                Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == i);
                values = new List<string>();
                for (int j = ColumnNameToNum(start.Col); j <= ColumnNameToNum(end.Col); j++)
                {
                    if (row == null) { values.Add(""); continue; }
                    Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == $"{ColumnNumToName(j)}{i}");
                    if (cell == null) { values.Add(""); continue; }
                    values.Add(GetCellValue(cell));
                }
                result.Add(values);
            }

            return result;
        }

        /// <summary>
        /// Performs a "Paste" action (overwriting existing cells) using the given "array" of strings, starting in the given cell
        /// </summary>
        /// <param name="worksheetName">The worksheet to paste into</param>
        /// <param name="start">A cell reference (eg F3)</param>
        /// <param name="values">The values to paste</param>
        public void PasteRange(string worksheetName, CellRef start, List<List<string>> values)
        {
            int col = ColumnNameToNum(start.Col), startCol = col, row = (int)start.Row - 1;
            IEnumerable<Sheet> sheets = _workbookPart.Workbook.Descendants<Sheet>();
            WorksheetPart wspart;
            try
            {
                string relId = sheets.First(s => worksheetName.Equals(s.Name)).Id;
                wspart = (WorksheetPart)_workbookPart.GetPartById(relId);
            }
            catch (InvalidOperationException) { wspart = _workbookPart.GetPartsOfType<WorksheetPart>().First(); }
            foreach (List<string> r in values)
            {
                col = startCol;
                foreach (string s in r)
                {
                    Insert(col, row, s, wspart);
                    col++;
                }
                row++;
            }
        }

        /// <summary>
        /// Performs a "cut" action on the given range, returning the range cut
        /// </summary>
        /// <param name="worksheetName">The worksheet to cut from</param>
        /// <param name="start">The cell reference for the beginning of the cut range</param>
        /// <param name="end">The cell reference for the end of the cut range</param>
        /// <returns></returns>
        public List<List<string>> CutRange(string worksheetName, CellRef start, CellRef end)
        {
            List<List<string>> result = new List<List<string>>();
            IEnumerable<Sheet> sheets = _workbookPart.Workbook.Descendants<Sheet>();
            WorksheetPart wspart;
            try
            {
                string relId = sheets.First(s => worksheetName.Equals(s.Name)).Id;
                wspart = (WorksheetPart)_workbookPart.GetPartById(relId);
            }
            catch (InvalidOperationException) { wspart = _workbookPart.GetPartsOfType<WorksheetPart>().First(); }
            SheetData sheetData = wspart.Worksheet.GetFirstChild<SheetData>();

            List<string> values;
            for (uint i = start.Row; i <= end.Row; i++)
            {
                Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == i);
                values = new List<string>();
                for (int j = ColumnNameToNum(start.Col); j <= ColumnNameToNum(end.Col); j++)
                {
                    if (row == null) { values.Add(""); continue; }
                    Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == $"{ColumnNumToName(j)}{i}");
                    if (cell == null) { values.Add(""); continue; }
                    values.Add(GetCellValue(cell));
                    cell.Remove();
                }
                result.Add(values);
            }
            wspart.Worksheet.Save();
            return result;
        }
        #endregion

        #region Utils
        /// <summary>
        /// Converts an Excel column name into a 0-indexed column number (A=0, Z=25, XFD=16383)
        /// </summary>
        /// <param name="columnName">The Excel column name (A,Z,XFD)</param>
        /// <returns>The 0-indexed column number that corresponds with this column name in Excel</returns>
        /// <remarks>Credits: http://stackoverflow.com/questions/667802/what-is-the-algorithm-to-convert-an-excel-column-letter-into-its-number </remarks>
        public static int ColumnNameToNum(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");
            columnName = columnName.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }
            return sum - 1;
        }

        /// <summary>
        /// Gets the Excel column name from a 0-indexed column number (0=A; 16383=XFD)
        /// </summary>
        /// <param name="columnNumber">The 0-indexed column number to convert</param>
        /// <returns>The Excel column name corresponding with the given 0-indexed column number</returns>
        /// <remarks>Credits: http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa </remarks>
        public static string ColumnNumToName(int columnNumber)
        {
            int dividend = columnNumber + 1; //the alg uses 1-indexed; I want 0
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        /// <summary>
        /// Gets a string representation of the information inside the given CellRef
        /// </summary>
        /// <param name="worksheetName">The worksheet to search</param>
        /// <param name="cellRef">The CellRef to check</param>
        /// <returns>The text inside the Cell</returns>
        public string GetCellValue(string worksheetName, CellRef cellRef)
        {
            string result;
            IEnumerable<Sheet> sheets = _workbookPart.Workbook.Descendants<Sheet>();
            WorksheetPart wspart;
            try
            {
                string relId = sheets.First(s => worksheetName.Equals(s.Name)).Id;
                wspart = (WorksheetPart)_workbookPart.GetPartById(relId);
            }
            catch (InvalidOperationException) { wspart = _workbookPart.GetPartsOfType<WorksheetPart>().First(); }
            SheetData sheetData = wspart.Worksheet.GetFirstChild<SheetData>();
            
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == cellRef.Row);            

            if (row == null) { result = ""; }
            else
            {
                Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellRef);
                if (cell == null) { result = ""; }
                else result = GetCellValue(cell);
            }
            return result;
        }

        /// <summary>
        /// Gets a string representation of the information inside the given Cell
        /// </summary>
        /// <param name="cell">The Cell to check</param>
        /// <returns>The text inside the Cell</returns>
        private string GetCellValue(Cell cell)
        {
            var value = cell.InnerText;
            if (cell.DataType != null)
            {
                if (cell.CellFormula != null)
                {
                    if (cell.CellFormula.FormulaType == CellFormulaValues.Shared)
                    {
                        UInt32Value index = cell.CellFormula.SharedIndex;
                        SheetData sheetData = cell.Ancestors<Worksheet>().First().GetFirstChild<SheetData>(); //get the SheetData from the Worksheet that is in this Cell's list of ancestors
                        CellFormula formula = sheetData.Elements<CellFormula>().First(f => f.SharedIndex == index);
                        return formula.Text;//actually need to get the modified formula based on the difference from here to the root
                    }
                    else
                    {
                        return cell.CellFormula.Text;
                    }
                }
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:

                        // For shared strings, look up the value in the
                        // shared strings table.
                        var stringTable =
                            _workbookPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                        // If the shared string table is missing, something 
                        // is wrong. Return the index that is in
                        // the cell. Otherwise, look up the correct text in 
                        // the table.
                        if (stringTable != null)
                        {
                            value =
                                stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = bool.FalseString;
                                break;
                            default:
                                value = bool.TrueString;
                                break;
                        }
                        break;
                    case CellValues.Date:
                        value = DateTime.Parse(value).ToLongDateString();
                        break;
                    default: break;
                }
            }
            return value;
        }

        /// <summary>
        /// Class with implicit operators to/from string, which holds column and row values of an Excel cell reference
        /// </summary>
        public class CellRef
        {
            /// <summary>
            /// The cell reference's column name
            /// </summary>
            public string Col;
            /// <summary>
            /// The cell reference's row number
            /// </summary>
            public uint Row;
            /// <summary>
            /// Returns a formatted string indicating the column name and row number
            /// </summary>
            /// <param name="cellRef">The CellRef object to turn into a string</param>
            public static implicit operator string(CellRef cellRef) { return $"{cellRef.Col}{cellRef.Row}"; }
            /// <summary>
            /// Creates a new CellRef object from a string that has a column name and row number
            /// </summary>
            /// <param name="s">The name of the cell, in string form</param>
            public static implicit operator CellRef(string s)
            {
                if (Regex.IsMatch(s, @"[a-zA-Z]+\d+"))
                {
                    string col = Regex.Match(s, @"[a-zA-Z]+").Value;
                    string match = Regex.Match(s, @"\d+").Value;
                    uint row = uint.Parse(match);
                    return new CellRef() { Col = col, Row = row };
                }
                else return null;
            }
        }

        /// <summary>
        /// Enum with Excel document directions (left,right,up,down)
        /// </summary>
        public enum Direction {
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
            LEFT, RIGHT, UP, DOWN
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
        }
        
        /// <summary>
        /// Scans to the given <see cref="Direction"/> from the cell A1, stopping at the last non-empty cell it finds
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet to scan.  If not found, the method will scan the first sheet in the workbook.</param>
        /// <param name="direction">The <see cref="Direction"/> in which to scan</param>
        /// <returns>The last non-empty cell in the given <see cref="Direction"/> from cell A1</returns>
        public string ScanToEnd(string worksheetName, Direction direction) { return ScanToEnd(worksheetName, direction, "A1"); }
        
        /// <summary>
        /// Scans to the given <see cref="Direction"/> from the given cell, stopping at the last non-empty cell it finds
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet to scan.  If not found, the method will scan the first sheet in the workbook.</param>
        /// <param name="direction">The <see cref="Direction"/> in which to scan</param>
        /// <param name="start">The cell to start the scan from</param>
        /// <returns>The last non-empty cell in the given <see cref="Direction"/></returns>
        public CellRef ScanToEnd(string worksheetName, Direction direction, CellRef start)
        {
            string result = start;
            IEnumerable<Sheet> sheets = _workbookPart.Workbook.Descendants<Sheet>();
            WorksheetPart wspart;
            try
            {
                string relId = sheets.First(s => worksheetName.Equals(s.Name)).Id;
                wspart = (WorksheetPart)_workbookPart.GetPartById(relId);
            }
            catch (InvalidOperationException) { wspart = _workbookPart.GetPartsOfType<WorksheetPart>().First(); }
            SheetData sheetData = wspart.Worksheet.GetFirstChild<SheetData>();
            if (direction == Direction.LEFT || direction == Direction.RIGHT)
            {
                int colNum = ColumnNameToNum(start.Col);
                
                Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == start.Row);
                Cell cell = new Cell();
                while (cell != null && colNum >= 0)
                {
                    cell = row?.Elements<Cell>()?.FirstOrDefault(c => c.CellReference == $"{ColumnNumToName(colNum)}{start.Row}");
                    if (cell != null) { result = $"{ColumnNumToName(colNum)}{start.Row}"; }
                    colNum += (direction == Direction.RIGHT ? 1 : -1);
                }
            }
            else
            {
                int rowNum = (int)start.Row;
                Row row;
                Cell cell = new Cell();
                while (cell != null && rowNum > 0)
                {
                    row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowNum);

                    cell = row?.Elements<Cell>()?.FirstOrDefault(c => c.CellReference == $"{start.Col}{rowNum}");
                    if (cell != null) { result = $"{start.Col}{rowNum}"; }
                    rowNum += (direction == Direction.DOWN ? 1 : -1);
                }
            }
            return result;
        }

        /// <summary>
        /// Returns the cell reference that is <paramref name="number"/> cells away from <paramref name="start"/> in the <see cref="Direction"/> indicated
        /// </summary>
        /// <param name="direction">The <see cref="Direction"/> to scan</param>
        /// <param name="start">The cell reference to start from</param>
        /// <param name="number">The number of cells to scan</param>
        /// <returns>A cell reference the indicated number of cells in the indicated <see cref="Direction"/> away from startCell (minimum A1)</returns>
        public static CellRef Scan(Direction direction, CellRef start, int number)
        {
            string result = start;

            if (direction == Direction.LEFT || direction == Direction.RIGHT)
            {
                int colNum = ColumnNameToNum(start.Col) + (direction == Direction.RIGHT ? 1 : -1) * number;
                if (colNum < 0) colNum = 0;
                result = $"{ColumnNumToName(colNum)}{start.Row}";                
            }
            else
            {
                int rowNum = (int)start.Row + (direction == Direction.DOWN ? 1 : -1) * number;
                if (rowNum < 1) rowNum = 1;
                result = $"{start.Col}{rowNum}";
            }
            return result;
        }

        /// <summary>
        /// Returns the number of rows between the two given cell references
        /// </summary>
        /// <param name="startCell">The cell reference to use as the start</param>
        /// <param name="endCell">The cell reference to use as the end</param>
        /// <returns>The number of rows between the two given cell references</returns>
        public static int RowsBetween(CellRef startCell, CellRef endCell)
        {
            return (int)endCell.Row - (int)startCell.Row;
        }

        /// <summary>
        /// Returns the number of columns between the two given cell references
        /// </summary>
        /// <param name="startCell">The cell reference to use as the start</param>
        /// <param name="endCell">The cell reference to use as the end</param>
        /// <returns>The number of columns between the two given cell references</returns>
        public static int ColsBetween(CellRef startCell, CellRef endCell)
        {
            return ColumnNameToNum(endCell.Col) - ColumnNameToNum(startCell.Col);
        }

        #endregion

        #region Export
        /// <summary>
        /// Saves the current document but leaves it open
        /// </summary>
        public void Save() { _doc.Save(); }

        /// <summary>
        /// Saves and closes the current document
        /// </summary>
        public void SaveAndClose() { _doc.Close(); }

        /// <summary>
        /// Saves and closes the current document, then gets a <see cref="MemoryStream"/> for exporting.
        /// </summary>
        /// <returns><see cref="MemoryStream"/></returns>
        public MemoryStream GetStream()
        {
            _doc.Close();
            _stream.Flush();
            _stream.Position = 0;
            return _stream;
        }

        /// <summary>
        /// Saves the current document with a new name, and returns the new document
        /// </summary>
        /// <param name="filename">The name of the desired new document</param>
        public ExcelDocument SaveAs(string filename)
        {
            //Workbook contains Sheets-->Sheet
            //Sheet object contains SheetName
            //WorksheetPart contains Worksheet-->SheetData
            //SheetData object contains Row objects, which contain Cell objects.
            ExcelDocument newdocument = new ExcelDocument(filename);
            foreach (Sheet oldsheet in this._sheets.Descendants<Sheet>())
            {
                WorksheetPart newwspart = newdocument.InsertSheet(oldsheet.Name);
                SheetData oldSheetData = ((WorksheetPart)_workbookPart.GetPartById(oldsheet.Id)).Worksheet.GetFirstChild<SheetData>();                
                foreach (Cell oldcell in oldSheetData.Descendants<Cell>())
                {
                    Insert(oldcell.CellReference.Value, GetCellValue(oldcell), newwspart);
                }
            }

            //document.wbpart.Workbook.ReplaceChild(this.sheets, document.sheets);//replace new (empty) sheet data with data from the old document
            this._doc.Close();
            newdocument.Save();
            return newdocument;
        }
        #endregion
    }
}
