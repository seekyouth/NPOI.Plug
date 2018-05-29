using NPOI.Plug.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.Plug.Demo
{
    class Class1
    {
        /// <summary>
        /// 批量导入库存
        /// </summary>
        public void BatchImportStock()
        {
            //文件
            //ImportResult result = new ImportResult();
            //HttpFileCollection files = HttpContext.Current.Request.Files;
            //DateTime dateNow = DateTime.Now;
            //string fileName = files[0].FileName;
            //string nullAgentCode = "";
            //int count = 0;
            //string path = "/UpFiles/ExcelFiles";
            //if (!ValidateExeclFileType(fileName))
            //{
            //    result.IsSuccess = false;
            //    result.Message = "文件格式不正确";
            //}
            //DirectoryInfo directory = new DirectoryInfo(Server.MapPath(path));
            //if (!directory.Exists)//不存在
            //    directory.Create();
            //fileName = fileName.Replace(":", "_").Replace(" ", "_").Replace("\\", "_").Replace("/", "_");
            //fileName = "/Stock" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + fileName;
            //string savefile = Server.MapPath(path + fileName);
            //files[0].SaveAs(savefile);
            //FileStream file = new FileStream(savefile, FileMode.Open, FileAccess.Read);
            //StockExcelImport Import = new StockExcelImport();
            //try
            //{
            //    result = Import.ImportTemplate(file, fileName, 6);
            //    if (result.IsSuccess)
            //    {
            //        DataTable stockTable = (DataTable)result.ExtraInfo;
            //        List<Base_Stock> inputList = Utils.helper.DataMappingHelper.DataTable2Entities<Base_Stock>(stockTable);
            //        foreach (Base_Stock item in inputList)
            //        {
            //            ResultInfo apiResult = CoreAPI.QueryAgentInfo(item.BelongerCode, "");
            //            if (apiResult.IsSuccess)
            //            {
            //                Base_Stock Stock = _iStock_Server.Get(m => m.BelongerCode == item.BelongerCode && m.GoodsDetailsNo == item.GoodsDetailsNo);
            //                if (Stock == null)
            //                {
            //                    #region 添加库存
            //                    Stock = new Base_Stock();
            //                    Stock.AddTime = dateNow;
            //                    Stock.BelongerCode = item.BelongerCode;
            //                    Stock.Goods_Id = item.Goods_Id;
            //                    Stock.GoodsDetailsNo = item.GoodsDetailsNo;
            //                    Stock.Inventory = item.Inventory; //可用库存量等于库存量
            //                    Stock.UsableInventory = item.Inventory;
            //                    Stock.WHouse_Id = 1;
            //                    Stock.MaxInventory = 0;
            //                    Stock.MinInventory = 0;
            //                    Stock.OnOrders = 0;
            //                    Stock.PreOrders = 0;
            //                    Stock.State = 1;
            //                    Stock.Units = 0;
            //                    Stock.AddTime = dateNow;
            //                    Stock.AddPerson = CurrentManagerInfo.user_name;
            //                    Stock.UpdateTime = dateNow;
            //                    Stock.UpdatePerson = CurrentManagerInfo.user_name;
            //                    Stock.TeamName = item.TeamName;
            //                    Stock.BelongerName = item.BelongerName;
            //                    _iStock_Server.Add(Stock);
            //                    #endregion
            //                }
            //                else
            //                {
            //                    Stock.Inventory += item.Inventory;
            //                    Stock.UsableInventory += item.Inventory;
            //                    _iStock_Server.Update(m => m.BelongerCode == item.BelongerCode && m.GoodsDetailsNo == item.GoodsDetailsNo, m => new Base_Stock { Inventory = Stock.Inventory, UsableInventory = Stock.UsableInventory, TeamName = item.TeamName, UpdateTime = dateNow, UpdatePerson = CurrentManagerInfo.user_name });
            //                }
            //                #region 添加进出库明细
            //                Base_AccessStock accessStock = new Base_AccessStock();
            //                accessStock.AccessType = 0;
            //                accessStock.AddTime = dateNow;
            //                //accessStock.DocumentMakerId = CurrentUser.id;
            //                accessStock.BelongerCode = item.BelongerCode;
            //                accessStock.Managers = CurrentManagerInfo.user_name;
            //                accessStock.Remark = "预存";
            //                accessStock.State = StockState.Done.GetHashCode();
            //                Thread.Sleep(10);
            //                accessStock.StockCode = "AS" + item.BelongerCode + DateTime.Now.ToString("HHmmssfff");
            //                accessStock.StockType = StockType.Enter.GetHashCode();
            //                accessStock.BelongerName = item.BelongerName;
            //                accessStock.TeamName = item.TeamName;
            //                _iAccessStock_Server.Add(accessStock);
            //                #endregion

            //                #region 添加入库明细
            //                Base_AccessStockDetails accessStockDetails = new Base_AccessStockDetails();
            //                accessStockDetails.GoodsDetailsNo = item.GoodsDetailsNo;
            //                accessStockDetails.Remark = "后台管理人员批量导入";
            //                accessStockDetails.GoodsId = item.Goods_Id;
            //                accessStockDetails.AccessStockAmount = item.Inventory;
            //                accessStockDetails.AccessStockId = accessStock.Id;
            //                _iAccessStockDetails_Server.Add(accessStockDetails);
            //                #endregion
            //            }
            //            else
            //            {
            //                count++;
            //                nullAgentCode += item.BelongerCode + ",";
            //            }
            //        }
            //        result.Message = string.Format("成功导入{0}条，{1}", inputList.Count() - count, count == 0 ? "" : "未找到代理商编号，请手动核实" + nullAgentCode);
            //    }
            //    else
            //    {
            //        //设置错误模版http路径
            //        result.Message = Request.Url.Authority + result.Message;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    result.IsSuccess = false;
            //    result.Message = ex.Message;
            //}

            //HttpContext.Current.Response.Write(JsonHerper.ConvertToJson(result));
            //HttpContext.Current.Response.End();
        }



        /// <summary>
        /// 导出
        /// </summary>
        public void BtnOutPut2()
        {
            
            // 2.设置单元格抬头
            // key：实体对象属性名称，可通过反射获取值
            // value：Excel列的名称
            Dictionary<string, string> cellheader = new Dictionary<string, string> {
                    { "BelongerName","代理商名称"},
                    { "BelongerCode","代理商编号"},
                    { "TeamName","所属团队"},
                    { "GoodsName","商品名称"},
                    { "SpecText", "商品规格"},
                    { "WHouseName","所属仓库"},
                    { "Inventory","库存量"},
                    { "UsableInventory","可用库存"},
                    { "PickGoodsCount","提货中数量"},
                    { "PickGoodsDoneCount","提货完成量"},
                    { "AddTime","添加时间"},
                    { "AddPerson","添加人"},
                    { "UpdateTime","更新时间"},
                    { "UpdatePerson","更新人"}
                };
            // 3.进行Excel转换操作，并返回转换的文件下载链接
            string urlPath = "/" + ExeclOutPort.EntityListToExcel2007(cellheader, null, "仓库列表");
           new DownHelper().HttpDownload(urlPath);
        }
    }
}
