using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace excelXLL
{
    /// <summary>
    /// 根据6位数编码查询地址
    /// </summary>
    public class checkAreaNumbercs
    {
        /// <summary>
        /// 根据6位数编码查询地址
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public  string CheckNum(int number)
        {
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(Resource1.ResourceManager.GetObject("AreaNumber") as string);
            XmlNode nodecount = xmldoc.SelectSingleNode("//*[@id='" + number.ToString() + "']");
            if (nodecount != null)
            {
                if (number.ToString().Substring(2)=="0000")
                {
                    string c = nodecount.Attributes["name"].Value;
                    return c;
                }
                if (number.ToString().Substring(4) == "00")
                {
                    string a1 = nodecount.ParentNode.Attributes["name"].Value;
                    string a2 = nodecount.Attributes["name"].Value;
                    string a = a1 + a2;
                    return a;
                }
                string s1 = nodecount.ParentNode.ParentNode.Attributes["name"].Value;
                string s2 = nodecount.ParentNode.Attributes["name"].Value;
                string s3 = nodecount.Attributes["name"].Value;
                string s = s1 + s2 + s3;
                return s;
            }
            else
            {
                string idCity = number.ToString().Substring(0, 4) + "00";
                XmlNode nodecity = xmldoc.SelectSingleNode("//*[@id='" + idCity + "']");
                if (nodecity != null)
                {
                    string a1 = nodecity.ParentNode.Attributes["name"].Value;
                    string a2 = nodecity.Attributes["name"].Value;
                    string a = a1 + a2;
                    return a;
                }
                else
                {
                    string idProvinces = number.ToString().Substring(0, 2) + "0000";
                    XmlNode nodeProvinces = xmldoc.SelectSingleNode("//*[@id='" + idProvinces + "']");
                    if (nodeProvinces != null)
                    {
                        string c = nodeProvinces.Attributes["name"].Value;
                        return c;
                    }
                }
                return "未知地址！";
            }
        }
    }
}
