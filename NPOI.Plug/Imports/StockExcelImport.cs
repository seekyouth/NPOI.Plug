using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.ComponentModel;
using NPOI.Plug;
using System.ComponentModel.Composition;
using NPOI.Plug.Helper;

namespace ActivityPeptide.Web.TempFile
{
    /// <summary>
    /// 设备批量注册服务
    /// </summary>
    [Export(typeof(ExcelImport))]
    public class StockExcelImport : ExcelImport
    {
        /// <summary>
        /// 任务状态缓存字典
        /// </summary>
        private static Dictionary<string, string> GetStatusDict()
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            //任务状态类型下拉选择
            dict["运行"] = "0";
            dict["停止"] = "1";
            return dict;
        }

        /// <summary>
        ///下拉选项校验
        /// </summary>
        /// <param name="e">校验参数</param>
        /// <returns>错误信息</returns>
        private static string SelectVerify(ImportVerifyParam e, object extra)
        {
            string result = "";
            result = ExcelImportHelper.GetCellMsg(e.CellValue, e.ColName, 0, true);
            if (string.IsNullOrEmpty(result))
            {
                var dict = extra as Dictionary<string, string>;
                if (dict != null)
                {
                    if (!dict.ContainsKey(e.CellValue.ToString()))
                    {
                        result += e.ColName + "下拉选项" + e.CellValue + "不存在";
                    }
                }
            }
            return result;
        }

        /// <summary>
        ///Cron表达式校验
        /// </summary>
        /// <param name="e">校验参数</param>
        /// <returns>错误信息</returns>
        private static string CronVerify(ImportVerifyParam e, object extra)
        {
            string result = "";
            result = ExcelImportHelper.GetCellMsg(e.CellValue, e.ColName, 200, false, true);
            if (string.IsNullOrEmpty(result))
            {
                //if (!QuartzHelper.ValidExpression(e.CellValue.ToString()))
                //{
                //    result += "Cron表达式格式不正确";
                //}
            }
            return result;
        }

        /// <summary>
        ///Cron表达式校验
        /// </summary>
        /// <param name="e">校验参数</param>
        /// <returns>错误信息</returns>
        private static string CronVerify2(ImportVerifyParam e, object extra)
        {
            string result = "";
            if (e.ColName.Equals("库存量"))
            {
                if (Convert.ToInt32(e.CellValue)<=0)
              {
                  result += "库存量必须大于0";
              }
            }
            return result;
        }

        /// <summary>
        /// Excel字段映射
        /// </summary>
        private static Dictionary<string, ImportVerify> dictFields = new List<ImportVerify> {
            new ImportVerify{ ColumnName="商品编号",FieldName="Goods_Id",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="产品明细编号",FieldName="GoodsDetailsNo",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="代理商名称",FieldName="BelongerName",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="代理商编号",FieldName="BelongerCode",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="仓库编号",FieldName="WHouse_Id",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="库存量",FieldName="Inventory",VerifyFunc =CronVerify2},
            //new ImportVerify{ ColumnName="可用库存量",FieldName="UsableInventory",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="添加时间",FieldName="AddTime",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="添加人",FieldName="AddPerson",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="更新时间",FieldName="UpdateTime",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="更新人",FieldName="UpdatePerson",VerifyFunc =CronVerify},
            new ImportVerify{ ColumnName="所属团队",FieldName="TeamName",VerifyFunc =CronVerify},
            
        }.ToDictionary(e => e.ColumnName, e => e);


        /// <summary>
        /// 业务类型
        /// </summary>
        public override ExcelImportType Type
        {
            get
            {
                return ExcelImportType.Task;
            }
        }

        /// <summary>
        /// Excel字段映射及校验缓存
        /// </summary>
        /// <returns>字段映射</returns>
        public override Dictionary<string, ImportVerify> DictFields
        {
            get
            {
                return dictFields;
            }
        }

        /// <summary>
        ///返回对应的导出模版数据
        /// </summary>
        /// <param name="FilePath">模版的路径</param>
        /// <param name="s">响应流</param>
        /// <returns>模版MemoryStream</returns>
        public override void GetExportTemplate(string FilePath, Stream s)
        {
            //写入下拉框值 任务状态
            var sheet = NPOIHelper.GetFirstSheet(FilePath);
            string[] taskStatus = GetStatusDict().Keys.ToArray();
            int dataRowIndex = StartRowIndex + 1;
            NPOIHelper.SetHSSFValidation(sheet, taskStatus, dataRowIndex, 3);
            sheet.Workbook.Write(s);
        }

        /// <summary>
        /// 获取额外的校验所需信息
        /// </summary>
        /// <param name="listColumn">所有列名集合</param>
        /// <param name="dt">dt</param>
        /// <returns>额外信息</returns>
        /// <remarks>
        /// 例如导入excel中含有下拉框 导入时需要判断选项值是否还存在，可以通过该方法查询选项值
        /// </remarks>
        public override Dictionary<string, object> GetExtraInfo(List<string> listColumn, DataTable dt)
        {
            Dictionary<string, object> extraInfo = new Dictionary<string, object>();
            foreach (string name in listColumn)
            {
                switch (name)
                {
                    case "Status":
                        extraInfo[name] = GetStatusDict();
                        break;
                    default:
                        break;
                }
            }
            return extraInfo;
        }


        public override object SaveImportData(DataTable dt, Dictionary<string, object> extraInfo)
        {
            //DBSQL dbsql = new DBSQL();
            //using (TransactionScope ts = new TransactionScope(TransactionScopeOption.Required))
            //{
            //    try
            //    {
            //        dbsql.Open(); //打开Connection连接  
            //        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(dbsql.Connection))
            //        {
            //            bulkCopy.BatchSize = dt.Rows.Count;
            //            bulkCopy.DestinationTableName = "StudentInfo";
            //            bulkCopy.WriteToServer(dt);
            //        }
            //        ts.Complete();
            //    }
            //    catch (Exception ex)
            //    {
            //        return 0;
            //    }
            //}
            //return dt.Rows.Count;

            return null;
        }



        /// <summary>  
        /// 批量插入  
        /// </summary>  
        /// <typeparam name="T">泛型集合的类型</typeparam>  
        /// <param name="conn">连接对象</param>  
        /// <param name="tableName">将泛型集合插入到本地数据库表的表名</param>  
        /// <param name="list">要插入大泛型集合</param>  
        public static void BulkInsert<T>(SqlConnection conn, string tableName, IList<T> list)
        {
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
            {
                bulkCopy.BatchSize = list.Count;
                bulkCopy.DestinationTableName = tableName;
                DataTable table = new DataTable();
                PropertyDescriptor[] props = TypeDescriptor.GetProperties(typeof(T))
                    .Cast<PropertyDescriptor>()
                    .Where(propertyInfo => propertyInfo.PropertyType.Namespace.Equals("System"))
                    .ToArray();
                int Length = props.Length;
                foreach (PropertyDescriptor propertyInfo in props)
                {
                    if (!propertyInfo.Name.Equals("ts_rec_filesUrl") && !propertyInfo.Name.Equals("ts_durationMinute"))
                    {
                        bulkCopy.ColumnMappings.Add(propertyInfo.Name, propertyInfo.Name);
                        table.Columns.Add(propertyInfo.Name, Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType);
                    }
                    else
                    { Length--; }
                }
                object[] values = new object[Length];
                foreach (T item in list)
                {
                    for (int i = 0; i < Length; i++)
                    {
                        values[i] = props[i].GetValue(item);
                    }
                    table.Rows.Add(values);
                }
                bulkCopy.WriteToServer(table);
            }
        }
    }
}
