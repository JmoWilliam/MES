using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.Runtime.Serialization;

namespace TestClient
{
    public class Serializer
    {

        #region 以byte[]为参数

        /// <summary>
        /// 序列化对象
        /// </summary>
        /// <param name="obj">要序列化的对象</param>
        /// <returns>返回二进制</returns>
        public static byte[] Serialize(Object obj)
        {
            if (obj != null)
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                MemoryStream ms = new MemoryStream();
                byte[] b;
                binaryFormatter.Serialize(ms, obj);
                ms.Position = 0;
                b = new Byte[ms.Length];
                ms.Read(b, 0, b.Length);
                ms.Close();
                return b;
            }
            else
                return new byte[0];
        }

        /// <summary>
        /// 反序列化对象
        /// </summary>
        /// <param name="b">要反序列化的二进制</param>
        /// <returns>返回对象</returns>
        public static object Deserialize(byte[] b)
        {
            if (b == null || b.Length == 0)
                return null;
            else
            {

                object result = new object();
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                MemoryStream ms = new MemoryStream();
                try
                {
                    ms.Write(b, 0, b.Length);
                    ms.Position = 0;
                    result = binaryFormatter.Deserialize(ms);
                    ms.Close();
                }
                catch
                {
                    result = null;
                }
                finally
                {
                    ms.Close();
                }
                return result;
            }
        }

        #endregion

        #region 以String为参数

        /// <summary>
        /// 序列化对象
        /// </summary>
        /// <param name="obj">要序列化的对象</param>
        /// <returns>返回字符串</returns>
        public static string SerializeObject(object obj)
        {
            IFormatter formatter = new BinaryFormatter();
            string result = string.Empty;
            using (MemoryStream stream = new MemoryStream())
            {
                formatter.Serialize(stream, obj);

                byte[] b = new byte[stream.Length];
                b = stream.ToArray();
                result = Convert.ToBase64String(b);
                //stream.Flush();
                stream.Close();
            }
            return result;
        }

        /// <summary>
        /// 反序列化对象
        /// </summary>
        /// <param name="str">要反序列化的字符串</param>
        /// <returns>返回对象</returns>
        public static object DeserializeObject(string str)
        {
            IFormatter formatter = new BinaryFormatter();
            byte[] b = Convert.FromBase64String(str);
            object obj = null;
            using (Stream stream = new MemoryStream(b, 0, b.Length))
            {
                obj = formatter.Deserialize(stream);
                stream.Close();
            }
            return obj;
        }

        #endregion
    } 
}
