using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace GetDataInfo
{
    //常用正式表达式文本

 public  class CommonRegex
    {


         public Dictionary<string, string> DicRegexText()
         {
             Dictionary<string, string> kvDictonary = new Dictionary<string, string>();
             kvDictonary.Add("提取姓名", @"([^\x00-\xff]{2,4})");
             kvDictonary.Add("提取身份证号码", @"(\d{17,18}[xX]?)");
             kvDictonary.Add("提取所在地区编码", @"(\d{6})\d+");
             kvDictonary.Add("提取出生日期", @"\d{6}(\d{8})\d+");
             kvDictonary.Add("提取性别", @"\d{14}(\d{3})[\d+xX]");

            return kvDictonary;
         }
        public void bindCbox(ComboBox cbo)
        {
            Dictionary<string, string> cboDic = DicRegexText();
            BindingSource bs = new BindingSource();
            bs.DataSource = cboDic;
            cbo.DataSource = bs;
            cbo.ValueMember = "Value";//"Key" ;//实际值
            cbo.DisplayMember = "Key";//显示的值


        }
    }
}
