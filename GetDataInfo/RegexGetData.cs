using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GetDataInfo
{
    class RegexGetData
    {
        private string personName;
        private string personCardId;
        


        public RegexGetData() { }

        public RegexGetData(string strData)
        {
            this.personName = GetFirstData(strData, @"([^\x00-\xff]{2,4})");//匹配名字

            //匹配身份证号码 ^[1-9]\d{5}[1-9]\d{3}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}([0-9]|[xX]) //(\d{17,18}[xX]?)
            this.personCardId = GetFirstData(strData, @"(\d{17,18}[xX]?)");
        }

        /// <summary>
        /// 提取第一组匹配到的记录
        /// </summary>
        /// <param name="strData">数据源，即需要匹配的文本</param>
        /// <param name="strRegex">正则表达式</param>
        /// <returns>返回匹配到的第一条记录</returns>
        public  string GetFirstData(string strData, string strRegex)//([^\x00-\xff]{2,4})
        {
            string result;

            Regex rex = new Regex(strRegex);
            Match match = rex.Match(strData);
            if (match.Success)
            {
                result = match.Groups[1].Value;
                return result;  
            }
            else
            {
                return null;
            }

        }

        /// <summary>
        /// 获取匹配到的所有结果
        /// </summary>
        /// <param name="strData">数据源，即需要匹配的文本</param>
        /// <param name="strRegex">正则表达式</param>
        /// <returns>返回一个匹配结果的集合</returns>
        public  List<string> GetAllData(string strData, string strRegex)
        {
            
            List<string> result = new List<string>();

            MatchCollection match = Regex.Matches(strData, strRegex);

            foreach (Match m in match)
            {

                result.Add(m.Value);
            }


            return result;

        }

        /// <summary>
        /// 判断是否匹配到相关数据
        /// </summary>
        /// <param name="strData">数据源，即需要判断的文本</param>
        /// <param name="strRegex">正则表达式</param>
        /// <returns>返回判断结果</returns>
        public  bool IsMatchSuccess(string strData, string strRegex)
        {

            Regex rex = new Regex(strRegex);
            return rex.IsMatch(strData);

        }

        /// <summary>
        /// 把匹配到的字符替换成新的文本(省略参数则默认为替换成空)
        /// </summary>
        /// <param name="strData">数据源，即需要替换的文本</param>
        /// <param name="strRegex">正则表达式</param>
        /// <param name="strReplace">新的文本，省略参数则默认为替换成空</param>
        /// <returns></returns>
        public  string ReplaceString(string strData, string strRegex,string strReplace="")
        {
            string strResult = Regex.Replace(strData, strRegex, strReplace);
            return strResult;

        }


        public string PersonName
        {
            get { return personName; }
            set { personName = value; }
        }

        public string PersonCardId
        {
            get { return personCardId; }
            set { personCardId = value; }
        }


    }
}
