using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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
            this.WorkbookFromArray(array);
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

        public ExcelHelper(Dictionary<int, Dictionary<int, object>> dic)
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
        }

        private string GetCellString(int rowNum, int columnNum)
        {
            return this.Sheet.GetRow(rowNum)?.GetCell(columnNum)?.ToString() ?? "";
        }

        private void SetCellValue(int rowNum, int columnNum, object value)
        {
            if (this.Sheet.GetRow(rowNum) == null)
            {
                this.Sheet.CreateRow(rowNum);
            }
            if (this.Sheet.GetRow(rowNum).GetCell(columnNum) == null)
            {
                this.Sheet.GetRow(rowNum).CreateCell(columnNum, CellType.String);
            }
            this.Sheet.GetRow(rowNum).GetCell(columnNum).SetCellType(CellType.String);
            this.Sheet.GetRow(rowNum).GetCell(columnNum).SetCellValue(value?.ToString() ?? "");
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

        private void WorkbookFromArray(object[][] array)
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
                    row.CreateCell(j, CellType.String).SetCellValue(array[i][j]?.ToString() ?? "");
                }
            }
        }

        private void WorkbookFromDictionary(Dictionary<int, Dictionary<int, object>> dic)
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
                    row.CreateCell(j, CellType.String).SetCellValue(dic[i][j]?.ToString() ?? "");
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
                column.ColumnName = firstRowAsColumnName ? this.FirstRow.GetCell(i)?.ToString() ?? "" : i.ToString();
                dt.Columns.Add(column);
            }
            for (int i = startRow; i < endRow; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = startColumn; j < endColumn; j++)
                {
                    dr[j - startColumn] = this.GetCellString(i, j);
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
                    array[i - startRow][j - startColumn] = this.GetCellString(i, j);
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
                    dic[i].Add(j, this.GetCellString(i, j));
                }
            }
            return dic;
        }

        private string[] GetColumnRange(int columnNum, int startRow, int endRow)
        {
            CheckBegin(ref startRow, ref columnNum);
            string[] arr = new string[endRow - startRow];
            for (int i = startRow; i < endRow; i++)
            {
                arr[i - startRow] = this.GetCellString(i, columnNum);
            }
            return arr;
        }

        private string[][] GetColumnByIndexs(int startRow, int endRow, int[] indexs)
        {
            string[][] arr = new string[endRow - startRow][];
            for (int i = this.FirstRowNum; i < this.LastRowNum; i++)
            {
                arr[i] = new string[indexs.Length];
                for (int j = 0; j < indexs.Length; i++)
                {
                    arr[i - this.FirstRowNum][j] = this.GetCellString(i, indexs[j]);
                }
            }
            return arr;
        }

        private string[][] GetRowByIndexs(int startColumn, int endColumn, int[] indexs)
        {
            string[][] arr = new string[indexs.Length][];
            for (int i = 0; i < indexs.Length; i++)
            {
                arr[i] = this.WorkbookToArrray(indexs[i], indexs[i] + 1, startColumn, endColumn)[0];
            }
            return arr;
        }

        private void SetValueFromDataTable(DataTable dt, int startRow, int startColumn)
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
                    this.SetCellValue(i + startRow, j + startColumn, dt.Rows[i][j]);
                }
            }
            InitData();
        }

        private void SetValueFromArray(object[][] array, int startRow, int startColumn)
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
                    this.SetCellValue(i + startRow, j + startColumn, array[i][j]);
                }
            }
            InitData();
        }

        private void SetValueFromDictionary(Dictionary<int, Dictionary<int, object>> dic)
        {
            if (dic == null)
            {
                throw new Exception($"parameter {nameof(dic)} can not be null");
            }
            foreach (KeyValuePair<int, Dictionary<int, object>> outerPair in dic)
            {
                foreach (KeyValuePair<int, object> innerPair in outerPair.Value)
                {
                    this.SetCellValue(outerPair.Key, innerPair.Key, innerPair.Value);
                }
            }
            InitData();
        }

        private void SetColumnRange(object[] arr, int columnNum, int startRow)
        {
            if (arr == null)
            {
                throw new Exception($"parameter {nameof(arr)} can not be null");
            }
            CheckBegin(ref startRow, ref columnNum);
            for (int i = 0; i < arr.Length; i++)
            {
                this.SetCellValue(i + startRow, columnNum, arr[i]);
            }
        }

        private void SaveToDisk(string savePath, bool overwrite)
        {
            if (string.IsNullOrEmpty(savePath))
            {
                throw new Exception("parameter savePath can not be null");
            }
            if (File.Exists(savePath))
            {
                if (overwrite)
                {
                    Debug.Print(savePath);
                    File.Delete(savePath);
                }
                else
                {
                    throw new Exception($"this file \"{savePath}\" is already exists in the directory");
                }
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

        private void InsertRow(object[] arr)
        {
            if (arr == null)
            {
                throw new Exception($"parameter {nameof(arr)} can not be null");
            }
            this.Sheet.CreateRow(this.LastRowNum);
            for (int i = 0; i < arr.Length; i++)
            {
                this.Sheet.GetRow(this.LastRowNum).CreateCell(this.FirstColumnNum + i, CellType.String).SetCellValue(arr[i]?.ToString() ?? "");
            }
            this.LastRowNum++;
        }

        private void InsertColumn(object[] arr)
        {
            if (arr == null)
            {
                throw new Exception("parameter arr can not be null");
            }
            for (int i = 0; i < arr.Length; i++)
            {
                SetCellValue(this.FirstRowNum + i, this.LastColumnNum, arr[i]);
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
            return this.WorkbookToArrray(rowIndex, rowIndex + 1, this.FirstColumnNum, this.LastColumnNum)[0];
        }

        public string[] GetRow(int rowIndex, int startColumn)
        {
            return this.WorkbookToArrray(rowIndex, rowIndex + 1, startColumn, this.LastColumnNum)[0];
        }

        public string[] GetRow(int rowIndex, int startColumn, int endColumn)
        {
            return this.WorkbookToArrray(rowIndex, rowIndex + 1, startColumn, endColumn)[0];
        }

        public string[][] GetRows(int startColumn, int endColumn, int[] rowIndexs)
        {
            return this.GetRowByIndexs(startColumn, endColumn, rowIndexs);
        }

        public Dictionary<int, string> GetRowWithColumnIndex(int rowIndex)
        {
            return this.WorkbookToDictionary(rowIndex, rowIndex + 1, this.FirstColumnNum, this.LastColumnNum)[rowIndex];
        }

        public string[] GetColumn(int columnIndex)
        {
            return this.GetColumnRange(columnIndex, this.FirstRowNum, this.LastRowNum);
        }

        public string[] GetColumn(int columnIndex, int startRow)
        {
            return this.GetColumnRange(columnIndex, startRow, this.LastRowNum);
        }

        public string[] GetColumn(int columnIndex, int startRow, int endRow)
        {
            return this.GetColumnRange(columnIndex, startRow, endRow);
        }

        public string[][] GetColumns(int startRow, int endRow, int[] columnIndexs)
        {
            return this.GetColumnByIndexs(startRow, endRow, columnIndexs);
        }

        public string GetValue(int rowNum, int columnNum)
        {
            return this.GetCellString(rowNum, columnNum);
        }

        public void Update(Dictionary<int, Dictionary<int, object>> dic)
        {
            this.SetValueFromDictionary(dic);
        }

        public void Update(object[][] array)
        {
            this.SetValueFromArray(array, this.FirstRowNum, this.FirstColumnNum);
        }

        public void Update(object[][] array, int startRow, int startColumn)
        {
            this.SetValueFromArray(array, startRow, startColumn);
        }

        public void Update(DataTable dt)
        {
            this.SetValueFromDataTable(dt, this.FirstRowNum, this.FirstColumnNum);
        }

        public void Update(DataTable dt, int startRow, int startColumn)
        {
            this.SetValueFromDataTable(dt, startRow, startColumn);
        }

        public void Update(object[] arr, int rowNum)
        {
            this.SetValueFromArray(new object[1][] { arr }, rowNum, this.FirstColumnNum);
        }

        public void Update(object[] arr, int columnNum, bool isColumn)
        {
            this.SetColumnRange(arr, this.FirstRowNum, columnNum);
        }

        public void Update(object[] arr, int columnNum, int startRow, bool isColumn)
        {
            this.SetColumnRange(arr, startRow, columnNum);
        }

        public void Update(object value, int row, int column)
        {
            this.CheckBegin(ref row, ref column);
            this.SetCellValue(row, column, value);
        }

        public void AppendRow(object[] arr)
        {
            this.InsertRow(arr);
        }

        public void AppendColumn(string[] arr)
        {
            this.InsertColumn(arr);
        }

        public void Add(Dictionary<int, Dictionary<int, object>> dic)
        {
            this.SetValueFromDictionary(dic);
        }

        public void Add(object[][] array)
        {
            this.SetValueFromArray(array, this.FirstRowNum, this.FirstColumnNum);
        }

        public void Add(object[][] array, int startRow, int startColumn)
        {
            this.SetValueFromArray(array, startRow, startColumn);
        }

        public void Add(DataTable dt)
        {
            this.SetValueFromDataTable(dt, this.FirstRowNum, this.FirstColumnNum);
        }

        public void Add(DataTable dt, int startRow, int startColumn)
        {
            this.SetValueFromDataTable(dt, startRow, startColumn);
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
