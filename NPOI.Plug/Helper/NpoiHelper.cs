using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Text;
using System.Text.RegularExpressions;

namespace NPOI.Plug.Helper
{
    /// <summary>
    /// NPOI操作EXECL帮助类
    /// </summary>
    public static class NPOIHelper
    {
        /// <summary>
        /// EXECL最大列宽
        /// </summary>
        public static readonly int MAX_COLUMN_WIDTH = 100 * 256;

        /// <summary>
        /// 生成EXECL文件，通过读取DataTable和列头映射信息
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="ColumnInfoList">列字段映射信息</param>
        /// <returns>文件流</returns>
        public static MemoryStream Export(DataTable dt, List<ColumnInfo> ColumnInfoList)
        {
            if (dt == null || ColumnInfoList == null)
            {
                throw new ArgumentNullException();
            }
            int rowHeight = 20;
            //每个标签页最多行数
            int sheetRow = 65536;
            HSSFWorkbook workbook = new HSSFWorkbook();

            //文本样式
            ICellStyle centerStyle = workbook.CreateCellStyle();
            centerStyle.VerticalAlignment = VerticalAlignment.Center;
            centerStyle.Alignment = HorizontalAlignment.Center;

            ICellStyle leftStyle = workbook.CreateCellStyle();
            leftStyle.VerticalAlignment = VerticalAlignment.Center;
            leftStyle.Alignment = HorizontalAlignment.Left;

            ICellStyle rightStyle = workbook.CreateCellStyle();
            rightStyle.VerticalAlignment = VerticalAlignment.Center;
            rightStyle.Alignment = HorizontalAlignment.Right;

            //寻找列头和DataTable之间映射关系
            foreach (DataColumn col in dt.Columns)
            {
                ColumnInfo info = ColumnInfoList.FirstOrDefault<ColumnInfo>(e => e.Field.Equals(col.ColumnName, StringComparison.OrdinalIgnoreCase));
                if (info != null)
                {
                    switch (info.Align.ToLower())
                    {
                        case "left":
                            info.Style = leftStyle;
                            break;
                        case "center":
                            info.Style = centerStyle;
                            break;
                        case "right":
                            info.Style = rightStyle;
                            break;
                    }
                    info.IsMapDT = true;
                }
            }

            int sheetNum = (int)Math.Ceiling(dt.Rows.Count * 1.0 / 65536);
            //最多生成5个标签页的数据
            sheetNum = sheetNum > 3 ? 3 : (sheetNum == 0 ? 1 : sheetNum);
            ICell cell = null;
            object cellValue = null;
            for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++)
            {
                ISheet sheet = workbook.CreateSheet();
                sheet.CreateFreezePane(0, 1, 0, 1);
                //输出表头
                IRow headerRow = sheet.CreateRow(0);
                //设置行高
                headerRow.HeightInPoints = rowHeight;
                //首行样式
                ICellStyle HeaderStyle = workbook.CreateCellStyle();
                HeaderStyle.FillPattern = FillPattern.SolidForeground;
                HeaderStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                IFont font = workbook.CreateFont();
                font.Boldweight = short.MaxValue;
                HeaderStyle.SetFont(font);
                HeaderStyle.VerticalAlignment = VerticalAlignment.Center; ;
                HeaderStyle.Alignment = HorizontalAlignment.Center;
                //输出表头信息 并设置表头样式
                int i = 0;
                foreach (var data in ColumnInfoList)
                {
                    cell = headerRow.CreateCell(i);
                    cell.SetCellValue(data.Header.Trim());
                    cell.CellStyle = HeaderStyle;
                    i++;
                }
                //开始循环所有行
                int iRow = 1;
                int startRow = sheetIndex * (sheetRow - 1);
                int endRow = (sheetIndex + 1) * (sheetRow - 1);
                endRow = endRow <= dt.Rows.Count ? endRow : dt.Rows.Count;
                for (int rowIndex = startRow; rowIndex < endRow; rowIndex++)
                {
                    IRow row = sheet.CreateRow(iRow);
                    row.HeightInPoints = rowHeight;
                    i = 0;
                    foreach (var item in ColumnInfoList)
                    {
                        cell = row.CreateCell(i);
                        if (item.IsMapDT)
                        {
                            cellValue = dt.Rows[rowIndex][item.Field];
                            cell.SetCellValue(cellValue != DBNull.Value ? cellValue.ToString() : string.Empty);
                            cell.CellStyle = item.Style;
                        }
                        i++;
                    }
                    iRow++;
                }
                //自适应列宽度
                for (int j = 0; j < ColumnInfoList.Count; j++)
                {
                    sheet.AutoSizeColumn(j);
                    int width = sheet.GetColumnWidth(j) + 2560;
                    sheet.SetColumnWidth(j, width > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH : width);
                }
            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            return ms;
        }


        /// <summary>
        /// 获取第一个Sheet
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>Sheet</returns>
        public static ISheet GetFirstSheet(string filePath)
        {
            using (Stream stream = new MemoryStream(File.ReadAllBytes(filePath)))
            {
                IWorkbook workbook = new HSSFWorkbook(stream);
                if (workbook.NumberOfSheets > 0)
                {
                    return workbook.GetSheetAt(0);
                }
            }
            return null;
        }

        #region "Excel模版数据读取相关"

        /// <summary>
        /// 通过输入流初始化workbook
        /// </summary>
        /// <param name="ins">输入流</param>
        /// <returns>workbook对象</returns>
        public static IWorkbook InitWorkBook(Stream ins)
        {
            string filepath = ((FileStream)(ins)).Name;
            if (Path.GetExtension(filepath).ToLower() == ".xls".ToLower())
            {
                return new HSSFWorkbook(ins);
            }
            else
            {
                return new XSSFWorkbook(ins);
            }

        }

        /// <summary>
        /// 从excel第一个sheet中读取数据
        /// </summary>
        /// <param name="ins">输入流</param>
        /// <param name="headRowIndex">标题行索引 默认为第6行</param>
        /// <param name="fSheet">第一个sheet</param>
        /// <returns>DataTable</returns>
        public static DataTable GetDataFromExcel(Stream ins, out ISheet fSheet, int headRowIndex = 0)
        {
            IWorkbook workbook = InitWorkBook(ins);
            fSheet = null;
            DataTable dt = new DataTable();
            if (workbook.NumberOfSheets > 0)
            {
                fSheet = workbook.GetSheetAt(0);
                if (fSheet.LastRowNum < headRowIndex)
                {
                    throw new ArgumentException("Excel模版错误,标题行索引大于总行数");
                }
                //读取标题行
                IRow row = null;
                ICell cell = null;
                row = fSheet.GetRow(headRowIndex);
                object objColumnName = null;
                for (int i = 0, length = row.LastCellNum; i < length; i++)
                {
                    cell = row.GetCell(i);
                    if (cell == null)
                    {
                        continue;
                    }
                    objColumnName = GetCellVale(cell);
                    if (objColumnName != null)
                    {
                        dt.Columns.Add(objColumnName.ToString());
                    }
                    else
                    {
                        dt.Columns.Add("");
                    }
                }
                //读取数据行
                object[] entityValues = null;
                int columnCount = dt.Columns.Count;
                for (int i = headRowIndex + 1, length = fSheet.LastRowNum; i < length+1; i++)
                {
                    row = fSheet.GetRow(i);
                    if (row == null)
                    {
                        continue;
                    }
                    entityValues = new object[columnCount];
                    //用于判断是否为空行
                    bool isHasData = false;
                    int dataColumnLength = row.LastCellNum < columnCount ? row.LastCellNum : columnCount;
                    for (int j = 0; j < dataColumnLength; j++)
                    {
                        cell = row.GetCell(j);
                        if (cell == null)
                        {
                            continue;
                        }
                        entityValues[j] = GetCellVale(cell);
                        if (!isHasData && j < columnCount && entityValues[j] != null)
                        {
                            isHasData = true;
                        }
                    }
                    if (isHasData)
                    {
                        dt.Rows.Add(entityValues);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 设置excel模版错误信息
        /// </summary>
        /// <param name="sheet">数据标签</param>
        /// <param name="rowindex">错误信息显示行</param>
        /// <param name="msg">错误信息</param>
        public static void SetTemplateErrorMsg(ISheet sheet, int rowindex, string msg)
        {
            IRow row = sheet.GetRow(rowindex);
            row = sheet.CreateRow(rowindex);
            if (row != null && !string.IsNullOrEmpty(msg))
            {
                sheet.AddMergedRegion(new CellRangeAddress(rowindex, rowindex, 0, row.LastCellNum));

                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                }
                ICellStyle cellStyle = sheet.Workbook.CreateCellStyle();
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.Alignment = HorizontalAlignment.Left;
                IFont font = sheet.Workbook.CreateFont();
                font.FontHeightInPoints = 12;
                font.Color = HSSFColor.Red.Index;
                cellStyle.SetFont(font);
                cell.CellStyle = cellStyle;
                cell.SetCellValue(msg);
            }
        }

        /// <summary>
        /// 获取数据行的错误信息提示样式
        /// </summary>
        /// <returns>错误数据行样式</returns>
        public static ICellStyle GetErrorCellStyle(IWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Left;
            IFont font = wb.CreateFont();
            //font.FontHeightInPoints = 12;
            font.Color = HSSFColor.Red.Index;
            cellStyle.SetFont(font);
            return cellStyle;
        }

        /// <summary>
        /// 获取标题行的错误信息提示样式
        /// </summary>
        /// <returns>错误标题行样式</returns>
        public static ICellStyle GetErrorHeadCellStyle(IWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Center;
            IFont font = wb.CreateFont();
            font.Boldweight = short.MaxValue;
            font.Color = HSSFColor.Red.Index;
            cellStyle.SetFont(font);
            cellStyle.FillPattern = FillPattern.SolidForeground;
            return cellStyle;
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>单元格值</returns>
        private static object GetCellVale(ICell cell)
        {
            object obj = null;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        obj = cell.DateCellValue;
                    }
                    else
                    {
                        obj = cell.NumericCellValue;
                    }
                    break;
                case CellType.String:
                    obj = cell.StringCellValue;
                    break;
                case CellType.Boolean:
                    obj = cell.BooleanCellValue;
                    break;
                case CellType.Formula:
                    obj = cell.CellFormula;
                    break;

            }
            return obj;
        }
        #endregion

        #region "设置下拉选项"
        /// <summary>
        /// 设置某些列的值只能输入预制的数据,显示下拉框
        /// </summary>
        /// <param name="sheet">要设置的sheet</param>
        /// <param name="textlist">下拉框显示的内容</param>
        /// <param name="firstRow">开始行</param>
        /// <param name="firstCol">开始列</param>
        /// <returns>设置好的sheet</returns>
        public static ISheet SetHSSFValidation(ISheet sheet,
                string[] textlist, int firstRow, int firstCol)
        {
            return SetHSSFValidation(sheet, textlist, firstRow, sheet.LastRowNum, firstCol, firstCol);
        }

        /// <summary>
        /// 设置某些列的值只能输入预制的数据,显示下拉框
        /// </summary>
        /// <param name="sheet">要设置的sheet</param>
        /// <param name="textlist">下拉框显示的内容</param>
        /// <param name="firstRow">开始行</param>
        /// <param name="endRow">结束行</param>
        /// <param name="firstCol">开始列</param>
        /// <param name="endCol">结束列</param>
        /// <returns>设置好的sheet</returns>
        public static ISheet SetHSSFValidation(ISheet sheet,
                string[] textlist, int firstRow, int endRow, int firstCol,
                int endCol)
        {
            IWorkbook workbook = sheet.Workbook;
            if (endRow > sheet.LastRowNum)
            {
                endRow = sheet.LastRowNum;
            }
            ISheet hidden = null;

            string hiddenSheetName = "hidden" + sheet.SheetName;
            int hIndex = workbook.GetSheetIndex(hiddenSheetName);
            if (hIndex < 0)
            {
                hidden = workbook.CreateSheet(hiddenSheetName);
                workbook.SetSheetHidden(sheet.Workbook.NumberOfSheets - 1, SheetState.Hidden);
            }
            else
            {
                hidden = workbook.GetSheetAt(hIndex);
            }

            IRow row = null;
            ICell cell = null;
            for (int i = 0, length = textlist.Length; i < length; i++)
            {
                row = hidden.GetRow(i);
                if (row == null)
                {
                    row = hidden.CreateRow(i);
                }
                cell = row.GetCell(firstCol);
                if (cell == null)
                {
                    cell = row.CreateCell(firstCol);
                }
                cell.SetCellValue(textlist[i]);
            }

            // 加载下拉列表内容  
            string nameCellKey = hiddenSheetName + firstCol;
            IName namedCell = workbook.GetName(nameCellKey);
            if (namedCell == null)
            {
                namedCell = workbook.CreateName();
                namedCell.NameName = nameCellKey;
                namedCell.RefersToFormula = string.Format("{0}!${1}$1:${1}${2}", hiddenSheetName, NumberToChar(firstCol + 1), textlist.Length);
            }
            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint(nameCellKey);

            // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列  
            CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
            // 数据有效性对象  
            HSSFDataValidation validation = new HSSFDataValidation(regions, constraint);
            //// 取消弹出错误框
            //validation.ShowErrorBox = false;
            sheet.AddValidationData(validation);
            return sheet;
        }
        #endregion

        #region "私有方法"
        /// 
        /// 把1,2,3,...,35,36转换成A,B,C,...,Y,Z
        /// 
        /// 要转换成字母的数字（数字范围在闭区间[1,36]）
        /// 
        private static string NumberToChar(int number)
        {
            if (1 <= number && 36 >= number)
            {
                int num = number + 64;
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] btNumber = new byte[] { (byte)num };
                return asciiEncoding.GetString(btNumber);
            }
            return "A";
        }
        #endregion


        #region 从Excel取数据并记录到List集合里
        /// <summary>
        /// 从Excel取数据并记录到List集合里
        /// </summary>
        /// <param name="cellHeard">单元头的值和名称;</param>
        /// <param name="filePath">保存文件绝对路径</param>
        /// <param name="errorMsg">错误信息</param>
        /// <returns>转换后的List对象集合</returns>
        public static List<T> ExcelToEntityList<T>(Dictionary<string, string> cellHeard, string filePath, out StringBuilder errorMsg) where T : new()
        {
            List<T> enlist = new List<T>();
            errorMsg = new StringBuilder();
            try
            {
                return Excel2003ToEntityList<T>(cellHeard, filePath, out errorMsg);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 从Excel2003取数据并记录到List集合里
        /// </summary>
        /// <param name="cellHeard">单元头的Key和Value;</param>
        /// <param name="filePath">保存文件绝对路径</param>
        /// <param name="errorMsg">错误信息</param>
        /// <returns>转换好的List对象集合</returns>
        private static List<T> Excel2003ToEntityList<T>(Dictionary<string, string> cellHeard, string filePath, out StringBuilder errorMsg) where T : new()
        {
            IWorkbook workbook = null;
            errorMsg = new StringBuilder(); // 错误信息,Excel转换到实体对象时，会有格式的错误信息
            List<T> enlist = new List<T>(); // 转换后的集合
            List<string> keys = cellHeard.Keys.ToList(); // 要赋值的实体对象属性名称
            try
            {
                using (FileStream fs = File.OpenRead(filePath))
                {
                    if (Regex.IsMatch(filePath, ".xls$"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(0); // 获取此文件第一个Sheet页
                    for (int i = 1; i <= sheet.LastRowNum; i++) // 从1开始，第0行为单元头
                    {
                        // 1.判断当前行是否空行，若空行就不在进行读取下一行操作，结束Excel读取操作
                        if (sheet.GetRow(i) == null)
                        {
                            break;
                        }
                        T en = new T();
                        string errStr = ""; // 当前行转换时，是否有错误信息，格式为：第1行数据转换异常：XXX列；
                        for (int j = 0; j < keys.Count; j++)
                        {
                            // 2.若属性头的名称包含'.',就表示是子类里的属性，那么就要遍历子类，eg：UserEn.TrueName
                            if (keys[j].IndexOf(".") >= 0)
                            {
                                // 2.1解析子类属性
                                string[] properotyArray = keys[j].Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                                string subClassName = properotyArray[0]; // '.'前面的为子类的名称
                                string subClassProperotyName = properotyArray[1]; // '.'后面的为子类的属性名称
                                System.Reflection.PropertyInfo subClassInfo = en.GetType().GetProperty(subClassName); // 获取子类的类型
                                if (subClassInfo != null)
                                {
                                    // 2.1.1 获取子类的实例
                                    var subClassEn = en.GetType().GetProperty(subClassName).GetValue(en, null);
                                    // 2.1.2 根据属性名称获取子类里的属性信息
                                    System.Reflection.PropertyInfo properotyInfo = subClassInfo.PropertyType.GetProperty(subClassProperotyName);
                                    if (properotyInfo != null)
                                    {
                                        try
                                        {
                                            // Excel单元格的值转换为对象属性的值，若类型不对，记录出错信息
                                            properotyInfo.SetValue(subClassEn, GetExcelCellToProperty(properotyInfo.PropertyType, sheet.GetRow(i).GetCell(j)), null);
                                        }
                                        catch (Exception e)
                                        {
                                            if (errStr.Length == 0)
                                            {
                                                errStr = "第" + i + "行数据转换异常：";
                                            }
                                            errStr += cellHeard[keys[j]] + "列；";
                                        }

                                    }
                                }
                            }
                            else
                            {
                                // 3.给指定的属性赋值
                                System.Reflection.PropertyInfo properotyInfo = en.GetType().GetProperty(keys[j]);
                                if (properotyInfo != null)
                                {
                                    try
                                    {
                                        // Excel单元格的值转换为对象属性的值，若类型不对，记录出错信息
                                        properotyInfo.SetValue(en, GetExcelCellToProperty(properotyInfo.PropertyType, sheet.GetRow(i).GetCell(j)), null);
                                    }
                                    catch (Exception e)
                                    {
                                        if (errStr.Length == 0)
                                        {
                                            errStr = "第" + i + "行数据转换异常：";
                                        }
                                        errStr += cellHeard[keys[j]] + "列；";
                                    }
                                }
                            }
                        }
                        // 若有错误信息，就添加到错误信息里
                        if (errStr.Length > 0)
                        {
                            errorMsg.AppendLine(errStr);
                        }
                        enlist.Add(en);
                    }
                }
                return enlist;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 从Excel获取值传递到对象的属性里
        /// </summary>
        /// <param name="distanceType">目标对象类型</param>
        /// <param name="sourceCell">对象属性的值</param>
        private static Object GetExcelCellToProperty(Type distanceType, ICell sourceCell)
        {
            object rs = distanceType.IsValueType ? Activator.CreateInstance(distanceType) : null;

            // 1.判断传递的单元格是否为空
            if (sourceCell == null || string.IsNullOrEmpty(sourceCell.ToString()))
            {
                return rs;
            }

            // 2.Excel文本和数字单元格转换，在Excel里文本和数字是不能进行转换，所以这里预先存值
            object sourceValue = null;
            switch (sourceCell.CellType)
            {
                case CellType.Blank:
                    break;

                case CellType.Boolean:
                    break;

                case CellType.Error:
                    break;

                case CellType.Formula:
                    break;

                case CellType.Numeric: sourceValue = sourceCell.NumericCellValue;
                    break;

                case CellType.String: sourceValue = sourceCell.StringCellValue;
                    break;

                case CellType.Unknown:
                    break;

                default:
                    break;
            }

            string valueDataType = distanceType.Name;

            // 在这里进行特定类型的处理
            switch (valueDataType.ToLower()) // 以防出错，全部小写
            {
                case "string":
                    rs = sourceValue.ToString();
                    break;
                case "int":
                case "int16":
                case "int32":
                    rs = (int)Convert.ChangeType(sourceCell.NumericCellValue.ToString(), distanceType);
                    break;
                case "float":
                case "single":
                    rs = (float)Convert.ChangeType(sourceCell.NumericCellValue.ToString(), distanceType);
                    break;
                case "datetime":
                    rs = sourceCell.DateCellValue;
                    break;
                case "guid":
                    rs = (Guid)Convert.ChangeType(sourceCell.NumericCellValue.ToString(), distanceType);
                    return rs;
            }
            return rs;
        }
        
        #endregion



      
    }






    /// <summary>
    /// NPOI拓展方法
    /// </summary>
    public static class NPOIExtend
    {
        /// <summary>
        /// 获取RGB对应NPOI颜色值
        /// </summary>
        /// <param name="workbook">当前wb</param>
        /// <param name="R"></param>
        /// <param name="G"></param>
        /// <param name="B"></param>
        /// <returns></returns>
        public static short GetXLColour(this HSSFWorkbook workbook, int R, int G, int B)
        {
            short s = 0;
            HSSFPalette XlPalette = workbook.GetCustomPalette();
            HSSFColor XlColour = XlPalette.FindColor((byte)R, (byte)G, (byte)B);
            if (XlColour == null)
            {
                if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 255)
                {
                    if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 64)
                    {
                        //NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE = 64;
                        //NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE += 1;
                        XlColour = XlPalette.AddColor((byte)R, (byte)G, (byte)B);
                    }
                    else
                    {
                        XlColour = XlPalette.FindSimilarColor((byte)R, (byte)G, (byte)B);
                    }

                    s = XlColour.Indexed;
                }
            }
            else
            {
                s = XlColour.Indexed;
            }
            return s;
        }

        /// <summary>
        /// 冻结表格
        /// </summary>
        /// <param name="sheet">sheet</param>
        /// <param name="colCount">冻结的列数</param>
        /// <param name="rowCount">冻结的行数</param>
        /// <param name="startCol">右边区域可见的首列序号，从1开始计算</param>
        /// <param name="startRow">下边区域可见的首行序号，也是从1开始计算</param>
        /// <example>
        /// sheet1.CreateFreezePane(0, 1, 0, 1); 冻结首行
        /// sheet1.CreateFreezePane(1, 0, 1, 0);冻结首列
        /// </example>
        public static void FreezePane(this ISheet sheet, int colCount, int rowCount, int startCol, int startRow)
        {
            sheet.CreateFreezePane(colCount, rowCount, startCol, startRow);
        }
    }
}

