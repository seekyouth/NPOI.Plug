using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using System.Collections;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using log4net;
using System.Reflection;
using NPOI.Plug.Helper;

namespace NPOI.Plug
{
  public  class ExeclOutPort
    {
       /// <summary>
        /// 实体类集合导出到EXCLE2003
        /// </summary>
        /// <param name="cellHeard">单元头的Key和Value;</param>
        /// <param name="enList">数据源</param>
        /// <param name="sheetName">工作表名称</param>
        /// <returns>文件的下载地址</returns>
        public static string EntityListToExcel2003(Dictionary<string, string> cellHeard, IList enList, string sheetName)
        {
            try
            {
                string fileName = sheetName + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls"; // 文件名称
                string urlPath = "UpFiles/ExcelFiles/" + fileName; // 文件下载的URL地址，供给前台下载
                string filePath = HttpContext.Current.Server.MapPath("\\" + urlPath); // 文件路径
                // 1.检测是否存在文件夹，若不存在就建立个文件夹
                string directoryName = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directoryName))
                {
                    Directory.CreateDirectory(directoryName);
                }

                // 2.解析单元格头部，设置单元头的中文名称
                HSSFWorkbook workbook = new HSSFWorkbook(); // 工作簿
                ISheet sheet = workbook.CreateSheet(sheetName); // 工作表
                IRow row = sheet.CreateRow(0);
                List<string> keys = cellHeard.Keys.ToList();
                for (int i = 0; i < keys.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(cellHeard[keys[i]]); // 列名为Key的值
                }

                // 3.List对象的值赋值到Excel的单元格里
                int rowIndex = 1; // 从第二行开始赋值(第一行已设置为单元头)
                foreach (var en in enList)
                {
                    IRow rowTmp = sheet.CreateRow(rowIndex);
                    for (int i = 0; i < keys.Count; i++) // 根据指定的属性名称，获取对象指定属性的值
                    {
                        string cellValue = ""; // 单元格的值
                        object properotyValue = null; // 属性的值
                        System.Reflection.PropertyInfo properotyInfo = null; // 属性的信息

                        // 3.1 若属性头的名称包含'.',就表示是子类里的属性，那么就要遍历子类，eg：UserEn.UserName
                        if (keys[i].IndexOf(".") >= 0)
                        {
                            // 3.1.1 解析子类属性(这里只解析1层子类，多层子类未处理)
                            string[] properotyArray = keys[i].Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                            string subClassName = properotyArray[0]; // '.'前面的为子类的名称
                            string subClassProperotyName = properotyArray[1]; // '.'后面的为子类的属性名称
                            System.Reflection.PropertyInfo subClassInfo = en.GetType().GetProperty(subClassName); // 获取子类的类型
                            if (subClassInfo != null)
                            {
                                // 3.1.2 获取子类的实例
                                var subClassEn = en.GetType().GetProperty(subClassName).GetValue(en, null);
                                // 3.1.3 根据属性名称获取子类里的属性类型
                                properotyInfo = subClassInfo.PropertyType.GetProperty(subClassProperotyName);
                                if (properotyInfo != null)
                                {
                                    properotyValue = properotyInfo.GetValue(subClassEn, null); // 获取子类属性的值
                                }
                            }
                        }
                        else
                        {
                            // 3.2 若不是子类的属性，直接根据属性名称获取对象对应的属性
                            properotyInfo = en.GetType().GetProperty(keys[i]);
                            if (properotyInfo != null)
                            {
                                properotyValue = properotyInfo.GetValue(en, null);
                            }
                        }

                        // 3.3 属性值经过转换赋值给单元格值
                        if (properotyValue != null)
                        {
                            cellValue = properotyValue.ToString();
                            // 3.3.1 对时间初始值赋值为空
                            if (cellValue.Trim() == "0001/1/1 0:00:00" || cellValue.Trim() == "0001/1/1 23:59:59")
                            {
                                cellValue = "";
                            }
                        }

                        // 3.4 填充到Excel的单元格里
                        rowTmp.CreateCell(i).SetCellValue(cellValue);
                    }
                    rowIndex++;
                }

                // 4.生成文件
                FileStream file = new FileStream(filePath, FileMode.Create);
                workbook.Write(file);
                file.Close();

                // 5.返回下载路径
                return urlPath;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        /// <summary>
        /// 将DataTable数据导出到Excel文件中(xlsx)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static string EntityListToExcel2007(Dictionary<string, string> cellHeard, IList enList, string sheetName)
        {

            string fileName = sheetName + "-" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx"; // 文件名称
            string urlPath = "UpFiles/ExcelFiles/" + fileName; // 文件下载的URL地址，供给前台下载
            string filePath = HttpContext.Current.Server.MapPath("\\" + urlPath); // 文件路径

            // 检测是否存在文件夹，若不存在就建立个文件夹
            string directoryName = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directoryName))
            {
                Directory.CreateDirectory(directoryName);
            }

            //删除文件夹的旧内容
           DownHelper.DeleteFile(HttpContext.Current.Server.MapPath("\\" + "UpFiles/ExcelFiles/"), "", 6);

            //解析单元格头部，设置单元头的中文名称
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet(sheetName); // 工作表
            IRow row = sheet.CreateRow(0);
            List<string> keys = cellHeard.Keys.ToList();
            for (int i = 0; i < keys.Count; i++)
            {
                row.CreateCell(i).SetCellValue(cellHeard[keys[i]]); // 列名为Key的值
            }

            // List对象的值赋值到Excel的单元格里
            int rowIndex = 1; // 从第二行开始赋值(第一行已设置为单元头)
            foreach (var en in enList)
            {
                IRow rowTmp = sheet.CreateRow(rowIndex);
                for (int i = 0; i < keys.Count; i++) // 根据指定的属性名称，获取对象指定属性的值
                {
                    string cellValue = ""; // 单元格的值
                    object properotyValue = null; // 属性的值
                    System.Reflection.PropertyInfo properotyInfo = null; // 属性的信息

                    // 3.1 若属性头的名称包含'.',就表示是子类里的属性，那么就要遍历子类，eg：UserEn.UserName
                    if (keys[i].IndexOf(".") >= 0)
                    {
                        // 3.1.1 解析子类属性(这里只解析1层子类，多层子类未处理)
                        string[] properotyArray = keys[i].Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                        string subClassName = properotyArray[0]; // '.'前面的为子类的名称
                        string subClassProperotyName = properotyArray[1]; // '.'后面的为子类的属性名称
                        System.Reflection.PropertyInfo subClassInfo = en.GetType().GetProperty(subClassName); // 获取子类的类型
                        if (subClassInfo != null)
                        {
                            // 3.1.2 获取子类的实例
                            var subClassEn = en.GetType().GetProperty(subClassName).GetValue(en, null);
                            // 3.1.3 根据属性名称获取子类里的属性类型
                            properotyInfo = subClassInfo.PropertyType.GetProperty(subClassProperotyName);
                            if (properotyInfo != null)
                            {
                                properotyValue = properotyInfo.GetValue(subClassEn, null); // 获取子类属性的值
                            }
                        }
                    }
                    else
                    {
                        // 3.2 若不是子类的属性，直接根据属性名称获取对象对应的属性
                        properotyInfo = en.GetType().GetProperty(keys[i]);
                        if (properotyInfo != null)
                        {
                            properotyValue = properotyInfo.GetValue(en, null);
                        }
                    }

                    // 3.3 属性值经过转换赋值给单元格值
                    if (properotyValue != null)
                    {
                        cellValue = properotyValue.ToString();
                        // 3.3.1 对时间初始值赋值为空
                        if (cellValue.Trim() == "0001/1/1 0:00:00" || cellValue.Trim() == "0001/1/1 23:59:59")
                        {
                            cellValue = "";
                        }
                    }

                    // 3.4 填充到Excel的单元格里
                    rowTmp.CreateCell(i).SetCellValue(cellValue);
                }
                rowIndex++;
            }

            // 4.生成文件
            FileStream file = new FileStream(filePath, FileMode.Create);
            xssfworkbook.Write(file);
            file.Close();

            // 5.返回下载路径
            return urlPath;


            ////转为字节数组
            //MemoryStream stream = new MemoryStream();
            //xssfworkbook.Write(stream);
            //var buf = stream.ToArray();

            ////保存为Excel文件
            //using (FileStream fs = new FileStream(urlPath, FileMode.Create, FileAccess.Write))
            //{
            //    fs.Write(buf, 0, buf.Length);
            //    fs.Flush();
            //}
        }



       

    }
}
