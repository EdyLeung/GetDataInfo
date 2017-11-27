using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace GetDataInfo
{
    /// <summary>
    /// 把获取Office对象或一些常用的方法封装成一个类
    /// </summary>
   public class ExcelObjCls
    {
       //构造函数
       public ExcelObjCls(Excel.Application app)
	        {
	            this.App = app;
	        }
       #region 属性获取或设置
       /// <summary>
       /// 获取或设置Excel的Application
	    /// </summary>
        public Excel.Application App {get; set; }

       /// <summary>
        /// 获取当前工作簿对象
       /// </summary>
        public Excel.Workbook ActiveWorkbook
	    {
	            get { return App.ActiveWorkbook;}
	    }

        /// <summary>
        /// 获取当前工作表对象
        /// </summary>
        public Excel.Worksheet ActiveSheet
        {
               get { return App.ActiveSheet; }
        }

       /// <summary>
        /// 获取或设置是否刷新界面
       /// </summary>
        public bool ScreentUpdate
       {
            get { return App.ScreenUpdating; }
            set { App.ScreenUpdating = value; }
       }
       /// <summary>
        /// 获取或设置是否开启系统警告
       /// </summary>
        public bool DisplayAlerts
	    {
           get { return App.DisplayAlerts; }
            set { App.DisplayAlerts = value; }
	    }
       #endregion



        #region 操作工作表对象
        /// <summary>
        /// 判断工作簿中是否包含特定名称的工作表
        /// </summary>
        /// <param name="wkBk">需要判断的工作簿</param>
        /// <param name="shtName">判定的工作表名称</param>
        /// <returns>  True：包含，否则不包含</returns>
        public bool Bk_HasSht(Excel.Workbook wkBk,string shtName ) 
        {
            bool bRet = false;
            foreach (Excel.Worksheet wkSht in wkBk.Worksheets)
            {
                if (wkSht.Name==shtName)
                {
                    bRet = true;
                    return bRet;
                }  
            }
            return bRet;
        }
       /// <summary>
        /// 删除工作簿中特定名称的工作表
       /// </summary>
        /// <param name="wkBk">需要操作的工作簿</param>
        /// <param name="shtName">判定的工作表名称</param>
        public void Bk_DeleteSht(Excel.Workbook wkBk, string shtName)
        {
            if (Bk_HasSht(wkBk,shtName))
            {
                if (wkBk.Worksheets.Count !=1)
                {
                    DisplayAlerts = false;
                        wkBk.Worksheets[shtName].Delete();
                    DisplayAlerts = true;  
                }
                else
                {
                    wkBk.Worksheets.Add(After: wkBk.Worksheets[wkBk.Worksheets.Count]);

                    DisplayAlerts = false;
                        wkBk.Worksheets[shtName].Delete();
                    DisplayAlerts = true;  
                }
            }
        }

       /// <summary>
        /// 工作簿中初始化一个新的工作表
       /// </summary>
        /// <param name="wkBk">需要操作的工作簿</param>
        /// <param name="shtName">工作表名称</param>
        /// <returns>工作表对象</returns>
       public Excel.Worksheet  Bk_GetNewSht(Excel.Workbook wkBk, string shtName)
       {
           Bk_DeleteSht(wkBk, shtName);
           Excel.Worksheet wkSht;
           wkSht =wkBk.Worksheets.Add(After:wkBk.Worksheets[wkBk.Worksheets.Count]);
           wkSht.Name = shtName;           
           return wkSht;
       }
        #endregion



        #region 写入数据都工作表中


       /// <summary>
       /// 二维数组直接导出到工作表指定单元格里
       /// </summary>
       /// <typeparam name="T">类型</typeparam>
       /// <param name="wkRng">单元格</param>
       /// <param name="arrayData">二维数组</param>
       public void ExportDataInSht<T>( Excel.Range wkRng,   T[,] arrayData)
       {
           wkRng.get_Resize(arrayData.GetLength(0), arrayData.GetLength(1)).Value = arrayData; 

       }



        #endregion

    }
}
