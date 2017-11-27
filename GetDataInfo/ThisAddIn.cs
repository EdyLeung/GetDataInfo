using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ToolExcel = Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace GetDataInfo
{
    public partial class ThisAddIn
    {
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //ExcelObjCls ExcelObj = new ExcelObjCls(Globals.ThisAddIn.Application);
            //////Excel.Worksheet wrksht;
            //ToolExcel.Worksheet WS = Globals.Factory.GetVstoObject(ExcelObj.ActiveSheet);


            //ToolExcel.ControlCollection controls = (ToolExcel.ControlCollection)WS.OLEObjects();
            //foreach (ToolExcel.Controls.PictureBox item in controls)
            //{
            //    MessageBox.Show(item.Tag.ToString());
            //}
        }   

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
