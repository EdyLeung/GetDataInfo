using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Reflection;


namespace GetDataInfo
{
    public partial class Ribbon1
    {
        ExcelObjCls ExcelObj=null; //new ExcelObjCls(Globals.ThisAddIn.Application);

        private ListObject lo;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            ExcelObj = new ExcelObjCls(Globals.ThisAddIn.Application);
            
        }

        private void btnGetIDCardInfo_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            
            
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

               
            
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            FrmSetConfig fsc = new FrmSetConfig();
            fsc.ShowDialog();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //Excel.Shape sh;
            GetFileInfo fileInfo = new GetFileInfo();
            GetIdCardInfo getidinfo;
            List<string> strFile = fileInfo.GetFilePath(true);
            Dictionary<String, GetIdCardInfo> idcard = new Dictionary<string, GetIdCardInfo>();
            List<string> strTitle = new List<string>() { "姓名", "身份证号", "所在地区", "出生日期","性别", "年龄", "照片" };
            string strfilename;
            if (strFile != null)
            {
                
                foreach (string strId in strFile)
                {
                     strfilename = fileInfo.GetFileStyle(strId, 2);
                     getidinfo = new GetIdCardInfo(strfilename);
                     if (!idcard.ContainsKey(strId))
                     {
                         idcard.Add(strId, getidinfo);
                     }


                }


                ExcelObj.App.Range["a1"].get_Resize(1, strTitle.Count).Value = strTitle.ToArray();
                //app.Range["a2"].get_Resize(idCardArray.GetLength(0), idCardArray.GetLength(1)).Value = idCardArray;
                int i = 2;
                ExcelObj.ScreentUpdate = false;
                string header = "'";

                InsertPicToExcel inPicExcel = new InsertPicToExcel();



                foreach (KeyValuePair<string, GetIdCardInfo> kvp in idcard)
                {
                    ExcelObj.App.Cells[i, 1].value = kvp.Value.IdCardName;
                    ExcelObj.App.Cells[i, 2].value =header+ kvp.Value.IdCardNumber;
                    ExcelObj.App.Cells[i, 3].value = kvp.Value.IdCardArea;
                    ExcelObj.App.Cells[i, 4].value = kvp.Value.IdCardBirthday;
                    ExcelObj.App.Cells[i, 5].value = kvp.Value.IdCardSex;
                    ExcelObj.App.Cells[i, 6].value = kvp.Value.IdCardAge;
                    Excel.Range rng = ExcelObj.App.Cells[i, 7];

                    //inPicExcel.InsertPicToRange(ExcelObj.ActiveSheet, rng, kvp.Key.ToString());
                    inPicExcel.InsertPictureBox(ExcelObj.ActiveSheet, rng, kvp.Key.ToString(), kvp.Value.IdCardName);


                    inPicExcel.GetPicWidthHeight(ExcelObj.ActiveSheet, rng, kvp.Key.ToString());
                    inPicExcel.InsertPicToComment(ExcelObj.App.Cells[i, 1], kvp.Key.ToString(), kvp.Value.IdCardName);
                   
                   
                   

                    i++;

                }
                ExcelObj.App.Range[ ExcelObj.App.Range["a1"],ExcelObj.App.Range["f1"].End[Excel.XlDirection.xlUp]].EntireColumn.AutoFit();


                ExcelObj.ScreentUpdate =true;


            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {

        }
 

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
 
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {


        }


    }
}
