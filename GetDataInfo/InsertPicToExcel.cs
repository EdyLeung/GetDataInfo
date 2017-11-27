using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ToolExcel = Microsoft.Office.Tools.Excel;
using ExcelControls = Microsoft.Office.Tools.Excel.Controls;
using System.Drawing;
using System.Diagnostics;

namespace GetDataInfo
{
    class InsertPicToExcel
    {

        private int rngColWidth;
        private int rngRowHeight;
        private Excel.Shape sh;

        INIFileOperate iniInfo= new INIFileOperate();

        //构造函数
        public InsertPicToExcel()
        {
            string strPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\Config.ini";	//INI文件的物理地址
            this.rngColWidth = Convert.ToInt32(iniInfo.ReadString("PicSize", "RangeColWidth", "", strPath));
            this.rngRowHeight = Convert.ToInt32(iniInfo.ReadString("PicSize", "RangeRowHeight", "", strPath));
        }

        /// <summary>
        /// 插入图片到指定的单元格
        /// </summary>
        /// <param name="shtPic">工作表</param>
        /// <param name="rngPic">单元格</param>
        /// <param name="picPath">图片地址</param>
        public void  InsertPicToRange (Excel.Worksheet shtPic,Excel.Range rngPic,string picPath)
        {
            rngPic.ColumnWidth =  rngColWidth;
            rngPic.RowHeight =  rngRowHeight;
            //插入路径下的图片
            sh = shtPic.Shapes.AddPicture(
                picPath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoCTrue,
                rngPic.Left,
                rngPic.Top,
                0,
                0 
                );

           
            //int pWidth =(int) sh.Width;
            //int pHeight = (int)sh.Height;
            //float hwb = rngRowHeight / pHeight < rngColWidth / pWidth ? rngRowHeight / pHeight : rngColWidth / pWidth;//确定行宽比，这样才能等比缩放

            sh.Width =(float) rngPic.Width;
            sh.Height = (float)rngPic.Height;
            sh.Placement = Excel.XlPlacement.xlMoveAndSize;//设定图片随单元格的变化而变化


        }

        /// <summary>
        /// 图片插入到指定单元格批注中
        /// </summary>
        /// <param name="rngPic">单元格</param>
        /// <param name="picPath">图片地址</param>
        /// <param name="picName">图片名称</param>
        public void InsertPicToComment(Excel.Range rngPic,string picPath,string picName)
        {
            //rngPic.ColumnWidth = rngColWidth;
            //rngPic.RowHeight = rngRowHeight;

            rngPic.AddComment();    //创建批注       
            rngPic.Comment.Shape.TextFrame.Characters().Text = picName;//在批注中记录文件名
            rngPic.Comment.Shape.TextFrame.Characters().Font.Size = 10;//文件名字体大小
            rngPic.Comment.Shape.TextFrame.AutoSize=true;//调整批注大小
            rngPic.Comment.Shape.Fill.UserPicture(picPath);//填充图片
            
            //rngPic.Comment.Shape.Height = (float)rngPic.Height;//批注高度
            //rngPic.Comment.Shape.Width = (float)rngPic.Width;//批注宽度

            rngPic.Comment.Shape.Height = this.GetPicHeight;//批注高度
            rngPic.Comment.Shape.Width = this.GetPicWidth;//批注宽度
           
        }

        /// <summary>
        /// 获取图片的原始宽度和高度
        /// </summary>
        /// <param name="shtPic">工作表</param>
        /// <param name="rngPic">单元格</param>
        /// <param name="picPath">图片地址</param>
        public  void GetPicWidthHeight(Excel.Worksheet shtPic,Excel.Range rngPic,string picPath)
        {
           sh = shtPic.Shapes.AddPicture(
                picPath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoCTrue,
                rngPic.Left,
                rngPic.Top,
                -1,
                -1 
                );
           this.GetPicHeight = sh.Height;
           this.GetPicWidth = sh.Width;
           sh.Delete();
        }

        /// <summary>
        /// 插入VSTO winform中的PictrueBox控件来显示图片
        /// </summary>
        /// <param name="shtPic">工作表</param>
        /// <param name="rngPic">单元格</param>
        /// <param name="picPath">图片路径</param>
        /// <param name="picName">图片名字</param>
        public void InsertPictureBox(Excel.Worksheet shtPic, Excel.Range rngPic, string picPath, string picName)
        {
            rngPic.ColumnWidth = rngColWidth;
            rngPic.RowHeight = rngRowHeight;

            //转换为Vsto对象
            ToolExcel.Worksheet wrk = Globals.Factory.GetVstoObject(shtPic);
            
            //声明一个Excel宿主控件对象
            ExcelControls.PictureBox  picBox = new ExcelControls.PictureBox();
            if (!wrk.Controls.Contains(picName))
            {
                picBox.Image = Image.FromFile(picPath);//导入图片到控件中
                picBox.SizeMode = PictureBoxSizeMode.StretchImage;
                picBox.Tag = picName;
                picBox.Click += picBox_Click;

                //添加dot net 控件
                wrk.Controls.AddControl(picBox, rngPic, picName);
                //picBox.Placement = Excel.XlPlacement.xlMoveAndSize;//随单元格大小变化而变化
            }
        }

        void picBox_Click(object sender, EventArgs e)
        {
            ToolExcel.Controls.PictureBox pic =(ToolExcel.Controls.PictureBox) sender;
            MessageBox.Show(pic.Tag.ToString());
        }




        public float GetPicWidth { get; private set; }
        public float GetPicHeight { get;private set; }


        //public int RngColWidth
        //{
        //    get { return rngColWidth; }
        //    set { rngColWidth = value; }
        //}

        //public int RngRowHeight
        //{
        //    get { return rngRowHeight; }
        //    set { rngRowHeight = value; }
        //}

    }
}
