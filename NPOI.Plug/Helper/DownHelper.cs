using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;

namespace NPOI.Plug.Helper
{
  public  class DownHelper : System.Web.UI.Page
    {

        /// <summary>
        /// http下载文件
        /// </summary>
        /// <param name="url">下载文件地址</param>
        /// <param name="path">文件存放地址，包含文件名</param>
        /// <returns></returns>
        public  void HttpDownload(string fileUrl)
        {
            string fileName = fileUrl.Substring(fileUrl.LastIndexOf('/') + 1);// 文件名称
            //string urlPath = "UpFiles/RecFiles/" + fileName; // 文件下载的URL地址，供给前台下载
            string filePath = HttpContext.Current.Server.MapPath("\\" + fileUrl); // 文件路径
            //删除文件夹的旧内容
            DeleteFile(HttpContext.Current.Server.MapPath("\\" + "UpFiles/ExcelFiles/"), "", 6);
            //以字符流的形式下载文件
            FileStream fs = new FileStream(filePath, FileMode.Open);
            byte[] bytes = new byte[(int)fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();
            Response.ContentType = "application/octet-stream";
            //通知浏览器下载文件而不是打开
            Response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8));
            Response.BinaryWrite(bytes);
            Response.Flush();
            Response.End();
        }


        /// <summary>
        /// 删除文件夹下的文件
        /// </summary>
        /// <param name="dirname">目录</param>
        /// <param name="search">搜索字符串</param>
        /// <param name="time">时间超过多少小时的删除</param>
        public static void DeleteFile(string dirname, string search, int time)
        {
            System.IO.DirectoryInfo dir = new DirectoryInfo(dirname);
            System.IO.FileInfo[] files;
            if (search != string.Empty)
                files = dir.GetFiles(search);
            else
                files = dir.GetFiles();

            if (files != null)
            {
                for (int i = 0; i < files.Length; i++)
                {
                    try
                    {
                        if (time != -1)
                        {
                            TimeSpan ts = System.DateTime.Now.Subtract(files[i].CreationTime);
                            if (ts.TotalHours > time)
                                files[i].Delete();
                        }
                        else
                        {
                            files[i].Delete();
                        }
                    }
                    catch (Exception ex)
                    {
                        //做日志记录
                        StringBuilder ErrString = new StringBuilder();
                        ErrString.AppendLine();
                        ErrString.AppendFormat("发生时间：{0}", System.DateTime.Now.ToString());
                        ErrString.AppendLine();
                        ErrString.AppendFormat("异常信息:{0}", ex.Message);
                        ErrString.AppendLine();
                        ErrString.AppendFormat("错误源:{0}", ex.Source);
                        ErrString.AppendLine();
                        ErrString.AppendFormat("堆栈信息:{0}", ex.StackTrace);
                        ErrString.AppendLine();
                        ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
                        log.Error(ErrString.ToString());
                    }
                }
            }
        }
    }
}
