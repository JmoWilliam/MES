/***********************************
*    Description：异步TCP客户端
*    Mountpoint：这是一个单例脚本，无需挂载，直接调用即可
*    Date：2019.06.25
*    Version：unity版本2017.2.0f3
*    Author：LJF
***********************************/
using System.Collections;
using System.Collections.Generic;
//using UnityEngine;
using System.Net;
using System.Net.Sockets;
using System;
using System.Text;


namespace TestClient
{
    //规范命名、添加注释、合理封装、限制访问权限    
    public class AsyncTcpClient
    {
        private string ip1;
        private int port1;
        byte[] ReadBytes = new byte[1024 * 1024];
        //单例
        public static AsyncTcpClient Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new AsyncTcpClient();
                }
                return instance;
            }
        }
        private static AsyncTcpClient instance;

        System.Net.Sockets.TcpClient tcpClient;

        //连接服务器
        public void ConnectServer(string ip, int port)//填写服务端IP与端口
        {
            //Debuger.EnableSave = true;
            ip1 = ip;
            port1 = port;
            try
            {
                tcpClient = new System.Net.Sockets.TcpClient();//构造Socket
                tcpClient.BeginConnect(IPAddress.Parse(ip), port, Lianjie, null);//开始异步
            }
            catch (Exception e)
            {
                //Debug.Log(e.Message);
            }
        }

        //连接判断
      public  void Lianjie(IAsyncResult ar)
        {
            if (!tcpClient.Connected)
            {
                //Debug.Log("服务器未开启，尝试重连。。。。。。");
                tcpClient.BeginConnect(IPAddress.Parse(ip1), port1, Lianjie, null);
                //IAsyncResult rest = tcpClient.BeginConnect(IPAddress.Parse(ip1), port1, Lianjie, null);
                //bool scu= rest.AsyncWaitHandle.WaitOne(3000);
            }
            else
            {
                //Debug.Log("连接上了");
                tcpClient.EndConnect(ar);//结束异步连接
                tcpClient.GetStream().BeginRead(ReadBytes, 0, ReadBytes.Length, ReceiveCallBack, null);
            }
        }


        //接收消息
      public  void ReceiveCallBack(IAsyncResult ar)
        {
            try
            {
                int len = tcpClient.GetStream().EndRead(ar);//结束异步读取
                if (len > 0)
                {

                    string str = Encoding.UTF8.GetString(ReadBytes, 0, len);
                    str = Uri.UnescapeDataString(str);
                    //将接收到的消息写入日志
                    //Debuger.Log(string.Format("收到主机:{0}发来的消息|{1}", ip1, str));
                    //Debug.Log(str);
                    tcpClient.GetStream().BeginRead(ReadBytes, 0, ReadBytes.Length, ReceiveCallBack, null);
                }
                else
                {
                    tcpClient = null;
                    //Debug.Log("连接断开,尝试重连。。。。。。");
                    ConnectServer(ip1, port1);
                }
            }
            catch (Exception e)
            {
                //Debug.Log(e.Message);
            }
        }

        //发送消息
        public void SendMsg(string msg)
        {
            byte[] msgBytes = HexStringToByteArray(msg);
            tcpClient.GetStream().BeginWrite(msgBytes, 0, msgBytes.Length, (ar) => {
                tcpClient.GetStream().EndWrite(ar);//结束异步发送
            }, null);//开始异步发送
        }

        // 将16进制字符串转换成字节数组
        private static byte[] HexStringToByteArray(string hexString)
        {
            int len = hexString.Length;
            byte[] data = new byte[len / 2];
            for (int i = 0; i < len; i += 2)
            {
                data[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            }
            return data;
        }

        /// <summary>
        /// 断开连接
        /// </summary>
        public void Close()
        {
            if (tcpClient != null && tcpClient.Client.Connected)
                tcpClient.Close();
            if (!tcpClient.Client.Connected)
            {
                tcpClient.Close();//断开挂起的异步连接
            }
        }

    }
}
