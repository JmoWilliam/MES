using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Sockets;
using System.Net;
using System.Threading;

namespace TestClient
{
    public static class SocketClass
    {
        private static Socket clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        private static byte[] result = new byte[1024];
        public static void Init()
        {
            //设定服务器IP地址
            IPAddress ip = IPAddress.Parse("127.0.0.1");
            try
            {
                clientSocket.Connect(new IPEndPoint(ip, 6000)); //配置服务器IP与端口
                Console.WriteLine("连接服务器成功");
            }
            catch
            {
                Console.WriteLine("连接服务器失败，请按回车键退出！");
                return;
            }

            Thread receiveThread = new Thread(ReceiveMessage);
            receiveThread.Start();

        }
        /// <summary>
        /// 接收消息
        /// </summary>
        /// <param name="clientSocket"></param>
        private static void ReceiveMessage()
        {
            while (true)
            {
                try
                {
                    //通过clientSocket接收数据
                    int receiveNumber = clientSocket.Receive(result);
                    string strContent = Encoding.ASCII.GetString(result, 0, receiveNumber);
                    Console.WriteLine("接收服务端{0}消息{1}", clientSocket.RemoteEndPoint.ToString(), Encoding.ASCII.GetString(result, 0, receiveNumber));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    clientSocket.Shutdown(SocketShutdown.Both);
                    clientSocket.Close();
                    break;
                }
            }
        }
        public static void SendMessage(string message)
        {
            string sendMessage = message;
            clientSocket.Send(Encoding.ASCII.GetBytes(sendMessage));
        }
    }
}
