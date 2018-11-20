using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CommonHelper
{
    public class ModelConvertHelper<T> where T : class, new()
    {
        /// <summary>
        /// 当对象属性为英文字段时
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static IList<T> DataTableToList(DataTable dt)
        {
            // 定义集合    
            IList<T> ts = new List<T>();

            // 获得此模型的类型   
            Type type = typeof(T);
            string tempName = "";
            foreach (DataRow dr in dt.Rows)
            {
                T t = new T();
                // 获得此模型的公共属性      
                PropertyInfo[] propertys = t.GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    tempName = pi.Name;  // 检查DataTable是否包含此列    

                    if (dt.Columns.Contains(tempName))
                    {
                        // 判断此属性是否有Setter      
                        if (!pi.CanWrite) continue;

                        object value = dr[tempName];
                        if (value != DBNull.Value)
                            pi.SetValue(t, value, null);
                    }
                }
                ts.Add(t);
            }
            return ts;
        }
        /// <summary>
        /// 当对象属性为中文字段时
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="xmlPath"></param>
        /// <returns></returns>
        public static IList<T> DataTableToList(DataTable dt, string xmlPath)
        {
            // 定义集合    
            IList<T> ts = new List<T>();

            // 获得此模型的类型   
            Type type = typeof(T);
            string tempName = "";
            var regulars = GetExportRegulars(xmlPath);
            foreach (DataRow dr in dt.Rows)
            {
                T t = new T();
                // 获得此模型的公共属性      
                PropertyInfo[] propertys = t.GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    // 检查DataTable是否包含此列    
                    var filedName = regulars.Where(p => p.PropertyName == pi.Name).Select(s => s.ExportFieldName).FirstOrDefault();
                    var filedType = regulars.Where(p => p.PropertyName == pi.Name).Select(s => s.DataType).FirstOrDefault();
                    if (dt.Columns.Contains(filedName))
                    {
                        // 判断此属性是否有Setter      
                        if (!pi.CanWrite) continue;
                        object value = dr[filedName];
                        switch (filedType)
                        {
                            case "DateTime":
                            case "Date":
                                value = Convert.ToDateTime(value);
                                break;
                            case "Time":
                                if (Convert.ToDateTime(value) == DateTime.MinValue)
                                {
                                    value = "";
                                }
                                else
                                {
                                    value = Convert.ToDateTime(value).ToString("HH:mm:ss");
                                }
                                break;
                            case "Int":
                                value = Convert.ToInt32(value);
                                break;
                            case "Double":
                                value = Convert.ToDouble(value);
                                break;
                            case "Decimal2":
                            case "Decimal4":
                            case "Decimal5":
                                value = Convert.ToDecimal(value);
                                break;
                            case "Bool":
                                value = "是";
                                if (!(bool)value)
                                {
                                    value = "否";
                                }
                                break;
                            default:
                                value = value.ToString();
                                break;
                        }
                        if (value != DBNull.Value)
                            pi.SetValue(t, value, null);
                    }
                }
                ts.Add(t);
            }
            return ts;
        }
        /// <summary>
        /// 解析XML规则集文件
        /// </summary>
        /// <returns></returns>
        public static List<ExportRegular> GetExportRegulars(string xmlpath)
        {
            var result = new List<ExportRegular>();

            var reader = new XmlTextReader(xmlpath);
            var doc = new XmlDocument();
            //从指定的XMLReader加载XML文档
            doc.Load(reader);

            foreach (XmlNode node in doc.DocumentElement.ChildNodes)
            {
                var header = new ExportRegular();

                if (node.Attributes["PropertyName"] != null)
                    header.PropertyName = node.Attributes["PropertyName"].Value;
                if (node.Attributes["DataType"] != null)
                    header.DataType = node.Attributes["DataType"].Value;
                if (node.Attributes["ExportFieldName"] != null)
                    header.ExportFieldName = node.Attributes["ExportFieldName"].Value;

                result.Add(header);
            }

            return result;
        }
    }
}
