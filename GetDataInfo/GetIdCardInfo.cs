using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetDataInfo
{
    public   class GetIdCardInfo
    {
        public string IdCardName{ get; set; }
        public string IdCardNumber { get; set; }
        public string IdCardArea { get; set; }
        public string IdCardBirthday { get; set; }
        public string IdCardSex { get; set; }
        public string IdCardAge { get; set; }
        public string FileName { get; set; }

        private RegexString regStr = new RegexString();

        private IdCardInfo   idcardinfo = new   IdCardInfo();

        public GetIdCardInfo(string  filename)
        {
            this.FileName = filename;

            GetIdInfo();

        }

        /// <summary>
        /// 获取身份证指定的信息
        /// </summary>
        public void GetIdInfo() 
        {
            this.IdCardName = idcardinfo.GetIdCardName(this.FileName, regStr.IdNameRegex);
            this.IdCardNumber = idcardinfo.GetIdCardNumber (this.FileName, regStr.IdCardRegex);
            this.IdCardArea = idcardinfo.GetArea (this.FileName, regStr.AreaRegex);
            this.IdCardBirthday = idcardinfo.GetBirthday (this.FileName, regStr.BrithRegex);
            this.IdCardSex = idcardinfo.GetSex(this.FileName, regStr.SexRegex);
            this.IdCardAge = idcardinfo.GetAgeByBirthdate(this.FileName, regStr.BrithRegex).ToString();
        }


    }
}
