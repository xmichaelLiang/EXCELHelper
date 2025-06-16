using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EXCELHelper
{
    public class ExcelNPOIService
    {
        
        /// <summary>
        /// 建立Excel，可選擇自訂Header/Body Style，未傳入則使用預設。
        /// 注意：headerStyle/bodyStyle 必須由同一個 IWorkbook 建立。
        /// </summary>
        public void CreateExcel<T>(
            string filePath,
            string sheetName,
            List<T> items)
        {
            IWorkbook book = CreateNewBook(filePath);
            ICellStyle headerStyle, bodyStyle;
                headerStyle = SetDefaultHeaderStyle(book);
                bodyStyle = SetDefaultBookBodyStyle(book);
            ISheet sheet = CreateEXCELSheet(sheetName, book);
            CreateSheetHeader(sheet, headerStyle, items);
            CreateBody(sheet, bodyStyle, items);
            SaveBook(book, filePath);
        }

        private void CreateBody<T>(ISheet sheet, ICellStyle bodyStyle, List<T> items)
        {
            int rowIndex = 1;
            var properties = typeof(T).GetProperties()
                .OrderBy(p => p.GetCustomAttribute<PropertySeqAttribute>(false)?.Seq ?? int.MaxValue)
                .ToArray();

            // 快取格式
            var styleCache = new Dictionary<Type, ICellStyle>();

            foreach (var item in items)
            {
                for (int colIndex = 0; colIndex < properties.Length; colIndex++)
                {
                    var value = properties[colIndex].GetValue(item);
                    var cellType = value?.GetType() ?? typeof(string);
                    var style = GetCellStyleForType(sheet.Workbook, bodyStyle, cellType, styleCache);

                    if (value == null)
                        WriteCell(sheet, colIndex, rowIndex, "", style);
                    else if (value is int)
                        WriteCell(sheet, colIndex, rowIndex, (int)value, style);
                    else if (value is double)
                        WriteCell(sheet, colIndex, rowIndex, (double)value, style);
                    else if (value is DateTime)
                        WriteCell(sheet, colIndex, rowIndex, (DateTime)value, style);
                    else
                        WriteCell(sheet, colIndex, rowIndex, value.ToString(), style);
                }
                rowIndex++;
            }

            // Auto size columns
            for (int colIndex = 0; colIndex < properties.Length; colIndex++)
            {
                sheet.AutoSizeColumn(colIndex);
            }
        }

        /// <summary>
        /// 根據型別取得對應的CellStyle，並快取
        /// </summary>
        private ICellStyle GetCellStyleForType(IWorkbook workbook, ICellStyle baseStyle, Type type, Dictionary<Type, ICellStyle> styleCache)
        {
            if (styleCache.TryGetValue(type, out var cachedStyle))
                return cachedStyle;

            ICellStyle style = workbook.CreateCellStyle();
            style.CloneStyleFrom(baseStyle);

            if (type == typeof(int) || type == typeof(int?))
                style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
            else if (type == typeof(double) || type == typeof(float) || type == typeof(decimal) ||
                     type == typeof(double?) || type == typeof(float?) || type == typeof(decimal?))
                style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            else if (type == typeof(DateTime) || type == typeof(DateTime?))
                style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
            else
                style.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");

            styleCache[type] = style;
            return style;
        }

        private ISheet CreateEXCELSheet(string SheetName,IWorkbook book) {

            ISheet sheet = book.CreateSheet(SheetName);
            return sheet;
        }


        private void CreateSheetHeader<T>(ISheet sheet, ICellStyle headerStyle, List<T> items)
        {
            int rowindex = 0;
            var properties = typeof(T)
                             .GetProperties()
                             .OrderBy(prop => prop.GetCustomAttribute<PropertySeqAttribute>(false)?.Seq ?? int.MaxValue).ToArray();
            for (int colIndex = 0; colIndex < properties.Length; colIndex++)
            {
                var columnName = properties[colIndex].GetCustomAttribute<PropertyColumnNameAttribute>(false)?.ColumnName??"欄位名稱未定義";
                WriteCell(sheet, colIndex, 0, columnName);
                WriteStyle(sheet, colIndex, rowindex, headerStyle);
                
            }
        }
     
        /// <summary>
        /// 設定Body Style（預設）
        /// </summary>
        public ICellStyle SetDefaultBookBodyStyle(IWorkbook book)
        {
            var style = book.CreateCellStyle();
            var font1 = book.CreateFont();
            //   font1.IsBold = true;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            // 不強制設為文字格式，讓 WriteCell 根據型別設定
            // style.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.BlueGrey.Index;
            style.SetFont(font1);
            return style;
        }

        /// <summary>
        /// 設定Header Style（預設）
        /// </summary>
        public ICellStyle SetDefaultHeaderStyle(IWorkbook book)
        {
            var style = book.CreateCellStyle();
            var font1 = book.CreateFont();
            font1.IsBold = true;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.BlueGrey.Index;
            style.SetFont(font1);
            return style;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fullFilePath"></param>
        /// <returns></returns>
        /// <exception cref="ApplicationException"></exception>
        private IWorkbook CreateNewBook(string fullFilePath)
        {
            IWorkbook book;
            var extension = Path.GetExtension(fullFilePath);

            // HSSF => Microsoft Excel(xls形式)(excel 97-2003)
            // XSSF => Office Open XML Workbook形式(xlsx形式)(excel 2007以降)
            if (extension == ".xls")
            {
                book = new HSSFWorkbook();
            }
            else if (extension == ".xlsx")
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            return book;
        }

        // 修改: WriteCell(string) 允許傳入格式
        private static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, string value, ICellStyle style = null)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
            if (style != null)
                cell.CellStyle = style;
        }

        // 修改: WriteCell(double) 允許傳入格式
        private void WriteCell(ISheet sheet, int columnIndex, int rowIndex, double value, ICellStyle style = null)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
            if (style != null)
                cell.CellStyle = style;
        }

        // 修改: WriteCell(DateTime) 允許傳入格式
        private void WriteCell(ISheet sheet, int columnIndex, int rowIndex, DateTime value, ICellStyle style = null)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
            if (style != null)
                cell.CellStyle = style;
        }

        //格式變更
        private void WriteStyle(ISheet sheet, int columnIndex, int rowIndex, ICellStyle style)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.CellStyle = style;
        }

        /// <summary>
        /// 儲存EXCEL檔
        /// </summary>
        /// <param name="book"></param>
        /// <param name="filePath"></param>
        private void SaveBook(IWorkbook book, string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                book.Write(fs);
            }
        }

        /// <summary>
        /// 解析Excel檔案，將資料轉換為List<T>，並根據[Required]屬性與型別進行檢核
        /// </summary>
        /// <typeparam name="T">目標Model型別</typeparam>
        /// <param name="filePath">Excel檔案路徑</param>
        /// <param name="sheetName">工作表名稱</param>
        /// <returns>List<T></returns>
        public List<T> ParseExcelToList<T>(string filePath, string sheetName) where T : new()
        {
            var workbook = LoadWorkbook(filePath);
            var sheet = GetSheet(workbook, sheetName);
            var headerDict = ReadHeaderRow(sheet);
            var propMap = MapPropertiesToColumns<T>(headerDict);
            var requiredProps = typeof(T).GetProperties()
                .Where(p => p.GetCustomAttribute<RequiredAttribute>() != null)
                .ToList();

            var result = new List<T>();
            for (int rowIdx = 1; rowIdx <= sheet.LastRowNum; rowIdx++)
            {
                var row = sheet.GetRow(rowIdx);
                if (row == null) continue;

                var instance = ParseRowToInstance<T>(row, propMap);
                ValidateRequiredProperties(instance, requiredProps, rowIdx + 1);
                result.Add(instance);
            }
            return result;
        }

        /// <summary>
        /// 讀取EXCEL檔案，返回IWorkbook物件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="ApplicationException"></exception>
        private IWorkbook LoadWorkbook(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var ext = Path.GetExtension(filePath);
                if (ext == ".xls")
                    return new HSSFWorkbook(fs);
                else if (ext == ".xlsx")
                    return new XSSFWorkbook(fs);
                else
                    throw new ApplicationException("LoadWorkbook: invalid extension");
            }
        }

        /// <summary>
        /// 取得指定名稱的工作表，如果不存在則返回第一個工作表
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private ISheet GetSheet(IWorkbook workbook, string sheetName)
        {
            var sheet = workbook.GetSheet(sheetName) ?? workbook.GetSheetAt(0);
            if (sheet == null) throw new Exception("Sheet not found.");
            return sheet;
        }

        /// <summary>
        /// 讀取工作表的第一行作為Header，並返回一個字典，鍵為欄位名稱，值為欄位索引
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private Dictionary<string, int> ReadHeaderRow(ISheet sheet)
        {
            var headerRow = sheet.GetRow(0);
            if (headerRow == null) throw new Exception("Header row not found.");

            int cellCount = headerRow.LastCellNum;
            var headerDict = new Dictionary<string, int>();
            for (int i = 0; i < cellCount; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                    headerDict[cell.ToString().Trim()] = i;
            }
            return headerDict;
        }

        /// <summary>
        /// 將Model的屬性映射到Excel的欄位，返回一個字典，鍵為欄位索引，值為屬性資訊
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="headerDict"></param>
        /// <returns></returns>
        private Dictionary<int, PropertyInfo> MapPropertiesToColumns<T>(Dictionary<string, int> headerDict)
        {
            var properties = typeof(T).GetProperties();
            var propMap = new Dictionary<int, PropertyInfo>();
            foreach (var prop in properties)
            {
                var colAttr = prop.GetCustomAttribute<PropertyColumnNameAttribute>(false);
                var colName = colAttr?.ColumnName ?? prop.Name;
                if (headerDict.TryGetValue(colName, out int colIdx))
                {
                    propMap[colIdx] = prop;
                }
            }
            return propMap;
        }

        /// <summary>
        /// 將Excel的行資料解析為Model實例，並根據屬性型別進行轉換
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <param name="propMap"></param>
        /// <returns></returns>
        private T ParseRowToInstance<T>(IRow row, Dictionary<int, PropertyInfo> propMap) where T : new()
        {
            var instance = new T();
            foreach (var kv in propMap)
            {
                var cell = row.GetCell(kv.Key);
                var prop = kv.Value;
                object value = null;

                if (cell != null)
                {
                    try
                    {
                        if (prop.PropertyType == typeof(string))
                            value = cell.ToString();
                        else if (prop.PropertyType == typeof(int) || prop.PropertyType == typeof(int?))
                            value = int.TryParse(cell.ToString(), out int i) ? (object)i : null;
                        else if (prop.PropertyType == typeof(double) || prop.PropertyType == typeof(double?))
                            value = double.TryParse(cell.ToString(), out double d) ? (object)d : null;
                        else if (prop.PropertyType == typeof(DateTime) || prop.PropertyType == typeof(DateTime?))
                        {
                            if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                                value = cell.DateCellValue;
                            else if (DateTime.TryParse(cell.ToString(), out DateTime dt))
                                value = dt;
                            else
                                value = null;
                        }
                        else
                            value = cell.ToString();
                    }
                    catch
                    {
                        value = null;
                    }
                }

                prop.SetValue(instance, value);
            }
            return instance;
        }

        /// <summary>
        /// 檢核Model的必填屬性，如果有未填寫的欄位則拋出ValidationException
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="instance"></param>
        /// <param name="requiredProps"></param>
        /// <param name="rowNumber"></param>
        /// <exception cref="ValidationException"></exception>
        private void ValidateRequiredProperties<T>(T instance, List<PropertyInfo> requiredProps, int rowNumber)
        {
            foreach (var reqProp in requiredProps)
            {
                var val = reqProp.GetValue(instance);
                if (val == null || (val is string s && string.IsNullOrWhiteSpace(s)))
                    throw new ValidationException($"Row {rowNumber}: 欄位「{reqProp.Name}」為必填。");
            }
        }
    }
}
