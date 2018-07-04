using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace QueryProduct
{
    public static class Product
    {
        private static readonly List<ProductInfo> List = new List<ProductInfo>();

        /// <summary>
        ///     根据编码读取价格
        /// </summary>
        /// <param name="arrList">配件编码</param>
        /// <param name="callback"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static string QueryProduct(ArrayList arrList,string filePath, Action<ProductInfo> callback)
        {
            try
            {
                foreach (object arr in arrList)
                {
                    if (arr == null) continue;

                    DataTable dt = ExcelHelper.QueryEx(filePath, arr.ToString());

                    foreach (DataRow row in dt.Rows)
                    {
                        string num = row[0].ToString();

                        var builder = new StringBuilder();
                        builder.AppendFormat("s_num={0}", num);

                        WebRequest request = WebRequest.Create("http://www.gwm.com.cn/CarFitting/Search");

                        byte[] buffer = Encoding.UTF8.GetBytes(builder.ToString());

                        request.ContentLength = buffer.Length;
                        request.Method = "POST";
                        request.ContentType = "application/x-www-form-urlencoded";

                        Stream requestStream = request.GetRequestStream();
                        requestStream.Write(buffer, 0, buffer.Length);
                        requestStream.Close();

                        WebResponse response = request.GetResponse();

                        Stream stream = response.GetResponseStream();

                        if (stream == null) return string.Empty;

                        var textReader = new StreamReader(stream);
                        var jsonReader = new JsonTextReader(textReader);
                        JToken array = JToken.ReadFrom(jsonReader);

                        var item = new ProductInfo();
                        if (array.Count() != 0)
                        {
                            item.Num = array[0]["Num"].ToString();
                            item.DPrice = array[0]["DPrice"].ToString();
                        }
                        else
                        {
                            item.Num = num;
                            item.DPrice = "AAA";
                        }

                        List.Add(item);

                        callback(item);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "查询出错啦~~");
            }

            return null;
        }

        /// <summary>
        /// 返回<see cref="ProductInfo"/>List对象
        /// </summary>
        /// <returns></returns>
        public static List<ProductInfo> QueryList()
        {
            return List;
        }
    }
}
