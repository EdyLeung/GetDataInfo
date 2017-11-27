using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDataInfo
{
 public   class RegexString
    {
     /// <summary>
     /// 正则表达式文本
     /// </summary>
        public string IdCardRegex { get; set; }
        public string IdNameRegex { get; set; }
        public string AreaRegex { get; set; }
        public string BrithRegex { get; set; }
        public string SexRegex { get; set; }

        private INIFileOperate iniInfo = new INIFileOperate();
        private string strPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\Config.ini";	//INI文件的物理地址
        public RegexString()
        {
            
            this.IdNameRegex = iniInfo.ReadString("RegString", "IdNameRegex", "", strPath);
            this.IdCardRegex = iniInfo.ReadString("RegString", "IdCardRegex", "", strPath);
            this.AreaRegex = iniInfo.ReadString("RegString", "AreaRegex", "", strPath);
            this.BrithRegex = iniInfo.ReadString("RegString", "BrithRegex", "", strPath);
            this.SexRegex = iniInfo.ReadString("RegString", "SexRegex", "", strPath);
  
        }
    }
}
