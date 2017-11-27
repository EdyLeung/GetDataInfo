using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace GetDataInfo
{
public  class IdCardInfo
    {
  
       private RegexGetData regdata = new RegexGetData();//创建一个正则表达式



        /// <summary>
        /// 获取身份证名字
        /// </summary>
        /// <param name="idCardString"></param>
        /// <param name="regexSting"></param>
        /// <returns></returns>
        public string GetIdCardName (string idCardString, string regexSting)
       {

           string IdCardName = regdata.GetFirstData(idCardString, regexSting);
           return IdCardName;
       }

        /// <summary>
        /// 获取身份证号码
        /// </summary>
        /// <param name="idCardString"></param>
        /// <param name="regexSting"></param>
        /// <returns></returns>
       public string GetIdCardNumber(string idCardString, string regexSting)
       {

        string IdCardNumber = regdata.GetFirstData(idCardString, regexSting);
        return IdCardNumber;
       }

 
    /// <summary>
    /// 获取性别
    /// </summary>
    /// <param name="idCardString"></param>
    /// <param name="regexSting"></param>
    /// <returns></returns>
        public string  GetSex(string idCardString,string  regexSting)
        {
            string sex = regdata.GetFirstData(idCardString, regexSting);

            if (int.Parse(sex) % 2 == 1)
            {
                return "男";
            }
            else
            {
                return "女";
            }
        }
    /// <summary>
    /// 获取所在地区
    /// </summary>
    /// <param name="idCardString"></param>
    /// <param name="regexSting"></param>
    /// <returns></returns>
        public string GetArea(string idCardString, string regexSting) 
        {

            string area = regdata.GetFirstData(idCardString, regexSting);



            string areaResult="";
            XmlDocument docXml = new XmlDocument();
            //docXml.Load(@"AreaCodeInfo.xml");

            docXml.Load("AreaCodeInfo.xml");
            if (getAreaName(docXml.DocumentElement, area) != null)
            {


                if (getAreaName(docXml.DocumentElement, area.Substring(0, 2).PadRight(6, '0')) != null)
                {
                    areaResult += getAreaName(docXml.DocumentElement, area.Substring(0, 2).PadRight(6, '0'));
                }

                if (getAreaName(docXml.DocumentElement, area.Substring(0, 4).PadRight(6, '0')) != null)
                {
                    areaResult += getAreaName(docXml.DocumentElement, area.Substring(0, 4).PadRight(6, '0'));
                }

                 areaResult += getAreaName(docXml.DocumentElement, area);

                 return areaResult;
            }
            else
            {
                return "区域代码错误！";
            }


        }
        /// <summary>
        /// 获取年龄
        /// </summary>
        /// <param name="birthdate">出生日期</param>
        /// <returns>返回年龄</returns>
        public int GetAgeByBirthdate(string idCardString, string regexSting) 
        {
            string strAge = regdata.GetFirstData(idCardString, regexSting);
            DateTime birthdate = DateTime.ParseExact(strAge, "yyyyMMdd", CultureInfo.CurrentCulture);

            DateTime now = DateTime.Now;
            int age = now.Year - birthdate.Year;
            if (now.Month < birthdate.Month || (now.Month == birthdate.Month && now.Day < birthdate.Day))
            {
                age--;
            }
            return age < 0 ? 0 : age;
        }

    /// <summary>
    /// 获取出生日期
    /// </summary>
    /// <param name="idCardString">字符串文本</param>
    /// <param name="regexSting">正则表达式</param>
    /// <returns></returns>
        public string GetBirthday(string idCardString, string regexSting) 
        {
            string age = regdata.GetFirstData(idCardString, regexSting);
            if (age  !="")
            {
                DateTime dateTime = DateTime.ParseExact(age, "yyyyMMdd", CultureInfo.CurrentCulture);
                return dateTime.ToLongDateString(); 
            }
            return null;
        }


    /// <summary>
    /// 转换为真正日期
    /// </summary>
    /// <param name="birthday"></param>
    /// <returns></returns>
        public DateTime ConvertDateTime(string birthday)
        {
      
                DateTime dateTime = DateTime.ParseExact(birthday, "yyyyMMdd", CultureInfo.CurrentCulture);
                return dateTime;     


        }

        private string getAreaName(XmlNode root, string areaCode)
        {
            string result = null;

            if (root == null)
            {
                return null;
            }

            if (root is XmlElement)
            {
                if (root.Attributes.Count > 0)
                {
                    if (root.Attributes["code"].Value.Equals(areaCode))
                    {
                        result = root.Attributes["name"].Value;
                        return result;
                    }
                }

                if (root.HasChildNodes)
                {
                    result = getAreaName(root.FirstChild, areaCode);
                }

                if (root.NextSibling != null)
                {
                    result = getAreaName(root.NextSibling, areaCode);
                }
            }

            return result;
        }







    }
}
