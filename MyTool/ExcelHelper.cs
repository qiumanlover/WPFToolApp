using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace MyTool
{
    public class ExcelHelper
    {
        public IWorkbook Workbook { get; private set; }
        private ISheet Sheet { get; set; }
        public int FirstRowNum { get; private set; }
        public int LastRowNum { get; private set; }
        public int FirstColumnNum { get; private set; }
        public int LastColumnNum { get; private set; }
        private IRow FirstRow { get; set; }
        private int SheetIndex { get; set; } = 0;


        public ExcelHelper(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new Exception($"parameter {nameof(filePath)} can not be null or empty");
            }
            string fileExt = Path.GetExtension(filePath);
            if (!fileExt.Equals(".xls") && !fileExt.Equals(".xlsx"))
            {
                throw new Exception($"the path {nameof(filePath)} is not a xls or xlsx file");
            }
            FileStream stream = null;
            try
            {
                using (stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    char ch = fileExt[fileExt.Length - 1];
                    if (ch.Equals('x'))
                    {
                        this.Workbook = new XSSFWorkbook(stream);
                    }
                    else
                    {
                        this.Workbook = new HSSFWorkbook(stream);
                    }
                }
            }
            catch (Exception)
            {
                throw new IOException($"open failed: can not open this file\r\n {filePath}\r\n this file may be locked or damaged");
            }
            finally
            {
                stream?.Dispose();
            }
            InitData();
        }

        public ExcelHelper(string[][] array)
        {
            if (array == null)
            {
                throw new Exception($"parameter {nameof(array)} can not be null");
            }
            this.WorkbookFromStringArray(array);
            InitData();
        }

        public ExcelHelper(DataTable table, bool columnNameAsFirstRow = true)
        {
            if (table == null)
            {
                throw new Exception($"parameter {nameof(table)} can not be null");
            }
            this.WorkbookFromDataTable(table, columnNameAsFirstRow);
            InitData();
        }

        public ExcelHelper(Dictionary<int, Dictionary<int, string>> dic)
        {
            if (dic == null)
            {
                throw new Exception($"parameter {nameof(dic)} can not be null");
            }
            this.WorkbookFromDictionary(dic);
            InitData();
        }

        private void InitData()
        {
            if ((this.Sheet = this.Workbook?.GetSheetAt(this.SheetIndex)) == null)
            {
                throw new Exception("excel file error: this file do not has a work sheet");
            }
            this.FirstRowNum = this.Sheet.FirstRowNum;
            this.LastRowNum = this.Sheet.LastRowNum;
            this.FirstRow = this.Sheet?.GetRow(this.FirstRowNum);
            this.FirstColumnNum = this.FirstRow?.FirstCellNum ?? -1;
            this.LastColumnNum = this.FirstRow?.LastCellNum ?? -1;
        }

        private void CheckRange(ref int startRow, ref int endRow, ref int startColumn, ref int endColumn)
        {
            startRow = startRow < this.FirstRowNum ? this.FirstRowNum : startRow;
            endRow = endRow > this.LastRowNum ? this.LastRowNum : endRow;
            startColumn = startColumn < this.FirstColumnNum ? this.FirstColumnNum : startColumn;
            endColumn = endColumn > this.LastColumnNum ? this.LastColumnNum : endColumn;
        }

        private void CheckBegin(ref int startRow, ref int startColumn)
        {
            startRow = startRow < this.FirstRowNum ? this.FirstRowNum : startRow;
            startColumn = startColumn < this.FirstColumnNum ? this.FirstColumnNum : startColumn;
            if (startRow >= this.LastRowNum || startColumn >= this.LastColumnNum)
            {
                throw new Exception($"start point ({startRow}, {startColumn}) is not in range");
            }
        }

        private void WorkbookFromDataTable(DataTable dt, bool columnNameAsFirstRow)
        {
            if (dt == null)
            {
                throw new Exception($"parameter {nameof(dt)} can not be null");
            }
            this.Workbook = new XSSFWorkbook();
            this.Sheet = this.Workbook.CreateSheet();
            int firstRowNum = columnNameAsFirstRow ? 1 : 0;
            if (columnNameAsFirstRow)
            {
                IRow firstRow = this.Sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    firstRow.CreateCell(i, CellType.String).SetCellValue(dt.Columns[i]?.ColumnName ?? "");
                }
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = this.Sheet.CreateRow(i + firstRowNum);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    row.CreateCell(j, CellType.String).SetCellValue(dt.Rows[i][j]?.ToString() ?? "");
                }
            }
        }

        private void WorkbookFromStringArray(string[][] array)
        {
            if (array == null)
            {
                throw new Exception($"parameter {nameof(array)} can not be null");
            }
            this.Workbook = new XSSFWorkbook();
            this.Sheet = this.Workbook.CreateSheet();
            for (int i = 0; i < array.Length; i++)
            {
                IRow row = this.Sheet.CreateRow(i);
                for (int j = 0; j < array[i].Length; j++)
                {
                    row.CreateCell(j, CellType.String).SetCellValue(array[i][j] ?? "");
                }
            }
        }

        private void WorkbookFromDictionary(Dictionary<int, Dictionary<int, string>> dic)
        {
            if (dic == null)
            {
                throw new Exception("parameter dic can not be null");
            }
            this.Workbook?.Close();
            this.Workbook = new XSSFWorkbook();
            this.Sheet = this.Workbook.CreateSheet();
            for (int i = 0; i < dic.Keys.Count; i++)
            {
                IRow row = this.Sheet.CreateRow(i);
                for (int j = 0; j < dic[i].Keys.Count; j++)
                {
                    row.CreateCell(j, CellType.String).SetCellValue(dic[i][j] ?? "");
                }
            }
        }

        private DataTable WorkbookToDataTable(int startRow, int endRow, int startColumn, int endColumn, bool firstRowAsColumnName)
        {
            if (endRow <= startRow || endColumn <= startColumn)
            {
                return null;
            }
            CheckRange(ref startRow, ref endRow, ref startColumn, ref endColumn);
            DataTable dt = new DataTable();
            for (int i = this.FirstColumnNum; i < this.LastColumnNum; i++)
            {
                DataColumn column = new DataColumn();
                column.DataType = Type.GetType("System.String");
                column.ColumnName = firstRowAsColumnName ? FirstRow.GetCell(i)?.ToString() ?? "" : i.ToString();
                dt.Columns.Add(column);
            }
            for (int i = startRow; i < endRow; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = startColumn; j < endColumn; j++)
                {
                    dr[j] = Sheet.GetRow(i)?.GetCell(j)?.ToString() ?? "";
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private string[][] WorkbookToArrray(int startRow, int endRow, int startColumn, int endColumn)
        {
            if (endRow <= startRow || endColumn <= startColumn)
            {
                return null;
            }
            CheckRange(ref startRow, ref endRow, ref startColumn, ref endColumn);
            string[][] array = new string[endRow - startRow][];
            for (int i = startRow; i < endRow; i++)
            {
                array[i] = new string[endColumn - startColumn];
                for (int j = startColumn; j < endColumn; j++)
                {
                    array[i][j] = Sheet.GetRow(i)?.GetCell(j)?.ToString() ?? "";
                }
            }
            return array;
        }

        private Dictionary<int, Dictionary<int, string>> WorkbookToDictionary(int startRow, int endRow, int startColumn, int endColumn)
        {
            if (endRow <= startRow || endColumn <= startColumn)
            {
                return null;
            }
            CheckRange(ref startRow, ref endRow, ref startColumn, ref endColumn);
            var dic = new Dictionary<int, Dictionary<int, string>>(endRow - startRow);
            for (int i = startRow; i < endRow; i++)
            {
                dic.Add(i, new Dictionary<int, string>(endColumn - startColumn));
                for (int j = startColumn; j < endColumn; j++)
                {
                    dic[i].Add(j, this.Sheet.GetRow(i)?.GetCell(j)?.ToString() ?? "");
                }
            }
            return dic;
        }

        private void UpdateFromDataTable(DataTable dt, int startRow, int startColumn)
        {
            if (dt == null)
            {
                throw new Exception($"parameter {nameof(dt)} can not be null");
            }
            CheckBegin(ref startRow, ref startColumn);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (this.Sheet.GetRow(i + startRow) == null)
                    {
                        this.Sheet.CreateRow(i + startRow);
                    }
                    if (this.Sheet.GetRow(i + startRow).GetCell(j + startColumn) == null)
                    {
                        this.Sheet.GetRow(i + startRow).CreateCell(j + startColumn, CellType.String);
                    }
                    this.Sheet.GetRow(i + startRow).GetCell(j + startColumn).SetCellType(CellType.String);
                    this.Sheet.GetRow(i + startRow).GetCell(j + startColumn).SetCellValue(dt.Rows[i][j]?.ToString() ?? "");
                }
            }
            InitData();
        }

        private void UpdateFromArray(string[][] array, int startRow, int startColumn)
        {
            if (array == null)
            {
                throw new Exception($"parameter {nameof(array)} can not be null");
            }
            CheckBegin(ref startRow, ref startColumn);
            for (int i = 0; i < array.Length; i++)
            {
                for (int j = 0; j < array[i].Length; j++)
                {
                    if (this.Sheet.GetRow(i + startRow) == null)
                    {
                        this.Sheet.CreateRow(i + startRow);
                    }
                    if (this.Sheet.GetRow(i + startRow).GetCell(j + startColumn) == null)
                    {
                        this.Sheet.GetRow(i + startRow).CreateCell(j + startColumn, CellType.String);
                    }
                    this.Sheet.GetRow(i + startRow).GetCell(j + startColumn).SetCellType(CellType.String);
                    this.Sheet.GetRow(i + startRow).GetCell(j + startColumn).SetCellValue(array[i][j] ?? "");
                }
            }
            InitData();
        }

        private void UpdateFromDictionary(Dictionary<int, Dictionary<int, string>> dic)
        {
            if (dic == null)
            {
                throw new Exception("parameter dic can not be null");
            }
            foreach (KeyValuePair<int, Dictionary<int, string>> outerPair in dic)
            {
                foreach (KeyValuePair<int, string> innerPair in outerPair.Value)
                {
                    if (this.Sheet.GetRow(outerPair.Key) == null)
                    {
                        this.Sheet.CreateRow(outerPair.Key);
                    }
                    if (this.Sheet.GetRow(outerPair.Key).GetCell(innerPair.Key) == null)
                    {
                        this.Sheet.GetRow(outerPair.Key).CreateCell(innerPair.Key, CellType.String);
                    }
                    this.Sheet.GetRow(outerPair.Key).GetCell(innerPair.Key).SetCellType(CellType.String);
                    this.Sheet.GetRow(outerPair.Key).GetCell(innerPair.Key).SetCellValue(innerPair.Value ?? "");
                }
            }
            InitData();
        }

        private string[] GetRowRange(int rowIndex, int startColumn, int endColumn)
        {
            string[] arr = new string[endColumn - startColumn];
            for (int i = startColumn; i < endColumn; i++)
            {
                arr[i] = this.Sheet.GetRow(rowIndex)?.GetCell(i)?.ToString() ?? "";
            }
            return arr;
        }

        private string[] GetColumnRange(int columnIndex, int startRow, int endRow)
        {
            string[] arr = new string[endRow - startRow];
            for (int i = startRow; i < endRow; i++)
            {
                arr[i] = this.Sheet.GetRow(i)?.GetCell(columnIndex)?.ToString() ?? "";
            }
            return arr;
        }

        private void SaveToDisk(string savePath, bool overwrite)
        {
            if (string.IsNullOrEmpty(savePath))
            {
                throw new Exception("parameter savePath can not be null");
            }
            if (overwrite)
            {
                File.Delete(savePath);
            }
            if (File.Exists(savePath))
            {
                throw new Exception($"this file \"{savePath}\" is already exists in the directory");
            }
            FileStream stream = null;
            try
            {
                using (stream = new FileStream(savePath, FileMode.Create, FileAccess.Write))
                {
                    this.Workbook.Write(stream);
                }
            }
            catch (Exception ex)
            {
                throw new IOException($"save file failed, can not save file in this directory: \r\n{Path.GetDirectoryName(savePath)}\r\n{ex.ToString()}");
            }
            finally
            {
                stream?.Dispose();
            }
        }

        private Stream SaveToStream()
        {
            MemoryStream ms = new MemoryStream();
            this.Workbook.Write(ms);
            return ms;
        }

        private string GetCellString(int rowNum, int columnNum)
        {
            return this.Sheet.GetRow(rowNum)?.GetCell(columnNum)?.ToString() ?? "";
        }

        private void UpdateCellValue(object value, int rowNum, int columnNum)
        {
            if (value == null)
            {
                throw new Exception("parameter value can not be null");
            }
            if (this.Sheet.GetRow(rowNum) == null)
            {
                this.Sheet.CreateRow(rowNum);
            }
            if (this.Sheet.GetRow(rowNum).GetCell(columnNum) == null)
            {
                this.Sheet.GetRow(rowNum).CreateCell(columnNum, CellType.String);
                InitData();
            }
            this.Sheet.GetRow(rowNum).GetCell(columnNum).SetCellType(CellType.String);
            this.Sheet.GetRow(rowNum).GetCell(columnNum).SetCellValue(value.ToString());
        }

        private void Insert(Dictionary<int, Dictionary<int, string>> dic)
        {
            if (dic == null)
            {
                throw new Exception("parameter dic can not be null");
            }
            foreach (KeyValuePair<int, Dictionary<int, string>> outerPair in dic)
            {
                if (this.Sheet.GetRow(outerPair.Key) == null)
                {
                    this.Sheet.CreateRow(outerPair.Key);
                }
                foreach (KeyValuePair<int, string> innerPair in outerPair.Value)
                {
                    if (this.Sheet.GetRow(outerPair.Key).GetCell(innerPair.Key) == null)
                    {
                        this.Sheet.GetRow(outerPair.Key).CreateCell(innerPair.Key, CellType.String).SetCellValue(innerPair.Value);
                    }
                    else
                    {
                        this.Sheet.GetRow(outerPair.Key).GetCell(innerPair.Key).SetCellType(CellType.String);
                        this.Sheet.GetRow(outerPair.Key).GetCell(innerPair.Key).SetCellValue(innerPair.Value);
                    }
                }
                InitData();
            }
        }

        private void InsertRow(string[] arr)
        {
            if (arr == null)
            {
                throw new Exception("parameter arr can not be null");
            }
            this.Sheet.CreateRow(this.LastRowNum);
            for (int i = 0; i < arr.Length; i++)
            {
                this.Sheet.GetRow(this.LastRowNum).CreateCell(this.FirstColumnNum + i, CellType.String).SetCellValue(arr[i]);
            }
            this.LastRowNum++;
        }

        private void InsertColumn(string[] arr)
        {
            if (arr == null)
            {
                throw new Exception("parameter arr can not be null");
            }
            for (int i = 0; i < arr.Length; i++)
            {
                this.Sheet.GetRow(this.FirstRowNum + i)?.CreateCell(this.LastColumnNum, CellType.String).SetCellValue(arr[i]);
            }
            this.LastColumnNum++;
        }

        private void SetCurSheet(int sheetIndex)
        {
            sheetIndex = sheetIndex >= this.Workbook.NumberOfSheets ? this.Workbook.NumberOfSheets - 1 : sheetIndex;
            sheetIndex = sheetIndex < 0 ? 0 : sheetIndex;
            this.SheetIndex = sheetIndex;
            InitData();
        }

        private void SetCurSheet(string sheetName)
        {
            if ((this.SheetIndex = this.Workbook.GetSheetIndex(sheetName)) < 0)
            {
                throw new Exception($"there is no sheet with name \"{nameof(sheetName)}\"");
            }
            InitData();
        }

        /*--------------------------------------------------------------------------------------------------------------------------------------------------*/

        public DataTable ToDataTable(bool firstRowAsColumnName = true)
        {
            return this.WorkbookToDataTable(this.FirstRowNum, this.LastRowNum, this.FirstColumnNum, this.LastColumnNum, firstRowAsColumnName);
        }

        public DataTable ToDataTable(bool withFirstRow, bool firstRowAsColumnName = true)
        {
            return withFirstRow ? this.ToDataTable() : this.WorkbookToDataTable(this.FirstRowNum + 1, this.LastRowNum, this.FirstColumnNum, this.LastColumnNum, firstRowAsColumnName);
        }

        public DataTable ToDataTable(int startRow, int endRow, int startColumn, int endColumn, bool firstRowAsColumnName = true)
        {
            return this.WorkbookToDataTable(startRow, endRow, startColumn, endColumn, firstRowAsColumnName);
        }

        public string[][] ToArray()
        {
            return this.WorkbookToArrray(this.FirstRowNum, this.LastRowNum, this.FirstColumnNum, this.LastColumnNum);
        }

        public string[][] ToArray(bool withFirstRow)
        {
            return withFirstRow ? this.ToArray() : this.WorkbookToArrray(this.FirstRowNum + 1, this.LastRowNum, this.FirstColumnNum, this.LastColumnNum);
        }

        public string[][] ToArray(int startRow, int endRow, int startColumn, int endColumn)
        {
            return this.WorkbookToArrray(startRow, endRow, startColumn, endColumn);
        }

        public Dictionary<int, Dictionary<int, string>> ToDictionary()
        {
            return this.WorkbookToDictionary(this.FirstRowNum, this.LastRowNum, this.FirstColumnNum, this.LastColumnNum);
        }

        public Dictionary<int, Dictionary<int, string>> ToDictionary(bool withFirstRow)
        {
            return withFirstRow ? this.ToDictionary() : this.WorkbookToDictionary(this.FirstRowNum + 1, this.LastRowNum, this.FirstColumnNum, this.LastColumnNum);
        }

        public Dictionary<int, Dictionary<int, string>> ToDictionary(int startRow, int endRow, int startColumn, int endColumn)
        {
            return this.WorkbookToDictionary(startRow, endRow, startColumn, endColumn);
        }

        public Stream ToStream()
        {
            return this.SaveToStream();
        }

        public string[] GetRow(int rowIndex)
        {
            return this.GetRowRange(rowIndex, this.FirstColumnNum, this.LastColumnNum);
        }

        public string[] GetRow(int rowIndex, int startColumn)
        {
            return this.GetRowRange(rowIndex, startColumn, this.LastColumnNum);
        }

        public string[] GetRow(int rowIndex, int startColumn, int endColumn)
        {
            return this.GetRowRange(rowIndex, startColumn, endColumn);
        }

        public string[] GetColumn(int columnIndex)
        {
            return this.GetRowRange(columnIndex, this.FirstRowNum, this.LastRowNum);
        }

        public string[] GetColumn(int columnIndex, int startRow)
        {
            return this.GetRowRange(columnIndex, startRow, this.LastRowNum);
        }

        public string[] GetColumn(int columnIndex, int startRow, int endRow)
        {
            return this.GetRowRange(columnIndex, startRow, endRow);
        }

        public string GetValue(int rowNum, int columnNum)
        {
            return this.GetCellString(rowNum, columnNum);
        }

        public void Update(Dictionary<int, Dictionary<int, string>> dic)
        {
            this.UpdateFromDictionary(dic);
        }

        public void Update(string[][] array)
        {
            this.UpdateFromArray(array, this.FirstRowNum, this.FirstColumnNum);
        }

        public void Update(string[][] array, int startRow, int startColumn)
        {
            this.UpdateFromArray(array, startRow, startColumn);
        }

        public void Update(DataTable dt)
        {
            this.UpdateFromDataTable(dt, this.FirstRowNum, this.FirstColumnNum);
        }

        public void Update(DataTable dt, int startRow, int startColumn)
        {
            this.UpdateFromDataTable(dt, startRow, startColumn);
        }

        public void Update(string[] arr, int rowNum)
        {
            this.UpdateFromArray(new string[1][] { arr }, rowNum, this.FirstColumnNum);
        }

        public void Update(string[] arr, int columnNum, bool isColumn = true)
        {
            this.UpdateFromArray(new string[1, 1] { arr }, this.FirstRowNum, columnNum);
        }

        public void Update(object value, int row, int column)
        {
            this.UpdateCellValue(value, row, column);
        }

        public void Add(string[] arr)
        {
            this.InsertRow(arr);
        }

        public void Add(string[] arr, bool toColumn = true)
        {
            this.InsertColumn(arr);
        }

        public void Add(Dictionary<int, Dictionary<int, string>> dic)
        {
            this.Insert(dic);
        }

        public void Save(string filePath, bool overwrite = false)
        {
            this.SaveToDisk(filePath, overwrite);
        }

        public void NextSheet()
        {
            this.SetCurSheet(++this.SheetIndex);
        }

        public void PreviousSheet()
        {
            this.SetCurSheet(--this.SheetIndex);
        }

        public void SetSheet(string sheetName)
        {
            this.SetCurSheet(sheetName);
        }
    }
}
