using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Threading;
using System.IO;
using System.Net;
using System.Timers;
//using UnityEngine;

namespace TestClient
{
    public partial class SocketFrm : Form
    {
        public SocketFrm()
        {
            InitializeComponent();
        }

       
        private void SocketFrm_Load(object sender, EventArgs e)
        {

        }


        public static bool IsConnet = true;//判断是否成功连接,设置为全局变量，方便随时控制
        //注意，这里的开始部分还是上一步的代码，只不过嵌进了方法体
        public void Connet(string Iptxt, int Port)//接收参数是目标ip地址和目标端口号。客户端无须关心本地端口号
        {
            //创建一个新的Socket对象
           // TcpClient 
            Socket client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IsConnet = true;//注意，此处是全局变量，将其设置为true
                    //将方法写进线程中
            Thread thread = new Thread(() =>
            {
                while (IsConnet)//循环
                {
                    try
                    {
                        client.Connect(Iptxt, Port);//尝试连接，失败则会跳去catch
                        IsConnet = false;//成功连接后修改bool值为false,这样下一步循环就不再执行。
                        break;//在此处加上break，成功就跳出循环，避免死循环
                    }
                    catch
                    {
                        client.Close();//先关闭
                        /*使用新的客户端资源覆盖，上一个已经废弃。如果继续使用以前的资源进行连接，
                        即使参数正确， 服务器全部打开也会无法连接*/
                        client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                        Thread.Sleep(1000);//等待1s再去重连
                    }
                }
                /*这里不一样就是放接收线程，在连接上后break出来，执行。
                因为需要带参数，所以要用到特别的ParameterizedThreadStart，
                然后开始线程。↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓*/
                Thread thread2 = new Thread(new ParameterizedThreadStart(ClientReceiveData));//接收线程方法
                thread2.IsBackground = true;//该值指示某个线程是否为后台线程。
                thread2.Start(client);//参数是用我们自建的Socket对象，就是上面的Socket client=new……

            });
            thread.IsBackground = true;//设置为后台线程，在程序退出时自己会自动释放
            thread.Start();//开始执行线程
        }
       



public void ClientReceiveData(object socket)//TCPClient消息的方法
        {
            var ProxSocket = socket as Socket;//处理上一步传过来的Socket函数
            byte[] data = new byte[1024 * 1024];//接收消息的缓冲区
            while (!IsConnet)//同样循环中止的条件
            {
                int len = 0;//记录消息长度，以及判断是否连接
                try
                {
                    //连接函数Receive会将数据放入data,从0开始放，之后返回数据长度。
                    len = ProxSocket.Receive(data, 0, data.Length, SocketFlags.None);
                }
                catch (Exception)
                {
                    //异常退出
                    ProxSocket.Shutdown(SocketShutdown.Both);//中止传输
                    ProxSocket.Close();//关闭
                    //Connet(tb_ServerIP.Text, 9004);//重新尝试去连接
                    IsConnet = false;//注意，此处是全局变量，将其设置为false,防止循环
                    return;//让方法结束，终结当前接收服务端数据的异步线程
                }
                if (len <= 0)
                {
                    //如果小于0，证明无连接，服务端正常退出
                    ProxSocket.Shutdown(SocketShutdown.Both);//中止传输
                    ProxSocket.Close();//关闭
                    //Connet(tb_ServerIP.Text, 9004);//重新尝试去连接
                    IsConnet = false;//注意，此处是全局变量，将其设置为false,防止循环
                    return;//让方法结束，终结当前接收服务端数据的异步线程
                }
                //这里做你想要对消息做的处理
                //string str = Encoding.Default.GetString(data, 0, len);//二进制数组转换成字符串……
            }
        }



        #region test code
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
         void Lianjie(IAsyncResult ar)
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
         void ReceiveCallBack(IAsyncResult ar)
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
                    StringBuilder responseMessage = new StringBuilder();
                    responseMessage.Append(Encoding.ASCII.GetString(ReadBytes, 0, ReadBytes.Length));
                    tb_RecieveMsg.Items.Add("Received response from server: /r/n" + responseMessage);
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

        #endregion

        private void btn_Link_Click(object sender, EventArgs e)
        {
            // 连接服务器


            //string serverAddress = tb_ServerIP.Text;     // 服务器IP地址
            //int serverPort =Convert.ToInt32(tb_ServerPort.Text);  // 服务器监听端口号

            //clientSocket.Connect(serverAddress, serverPort);
            //MessageBox.Show("登录服务器成功！\n");
        }


        //private AsyncTcpClient Socket1 = new AsyncTcpClient();
        private void btn_Send_Click(object sender, EventArgs e)
        {
            TcpClient clientSocket = new TcpClient();
            try
            {
                //创建TCP客户端套接字
                clientSocket.Connect(tb_ServerIP.Text, Convert.ToInt32(tb_ServerPort.Text));
                //Connet(tb_ServerIP.Text, Convert.ToInt32(tb_ServerPort.Text));
                if (clientSocket.Connected)
                {
                    tb_RecieveMsg.Items.Add("Server is Connected ！");
                    // 发送消息
                    string message = "4C4F4E0D";
                    byte[] data = HexStringToByteArray(message);
                    NetworkStream stream = clientSocket.GetStream();
                    stream.Write(data, 0, data.Length);

                    if (stream.ToString() != "")
                    {
                        // 接收服务器回复
                        tb_RecieveMsg.Items.Add("Start to Received ...");
                        byte[] responseData = new byte[256];
                        StringBuilder responseMessage = new StringBuilder();

                        int bytes = stream.Read(responseData, 0, responseData.Length);

                        responseMessage.Append(Encoding.ASCII.GetString(responseData, 0, bytes));
                        string[] dateString = responseMessage.ToString().Split(':');
                        string[] readString = dateString[1].ToString().Split(',');
                        //Console.WriteLine("Received response from server: " + responseMessage);
                        tb_RecieveMsg.Items.Add("Received response from server: " + responseMessage);
                        tb_RecieveMsg.Items.Add("Received response from server: " + dateString[0]);
                        tb_RecieveMsg.Items.Add("Data Count: " + readString.Count());
                        tb_RecieveMsg.Items.AddRange(readString);
                    }
                    else
                    {
                        throw new System.Exception("Can't Received Response !");
                    }
                    
                    // 关闭套接字
                    stream.Dispose();
                    stream.Close();
                }
            }
            catch(Exception ex)
            {
                tb_RecieveMsg.Items.Add(ex.Message);
            }
            clientSocket.Close();
        }

        private delegate void delegateUserControl(string str);
        private void ReceivedDataWrite(string receivedData)
        {
            tb_RecieveMsg.Items.Add(("[" + tb_ServerIP.Text + "][" + DateTime.Now + "]" + receivedData));
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

        private void tb_ServerIP_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_RecieveMsg_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
