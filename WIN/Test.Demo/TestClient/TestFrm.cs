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

namespace TestClient
{
    public partial class TestFrm : Form
    {
        public TestFrm()
        {
            InitializeComponent();
        }

        private void TestFrm_Load(object sender, EventArgs e)
        {

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
        public static bool IsConnet = true;//判断是否成功连接,设置为全局变量，方便随时控制
        //注意，这里的开始部分还是上一步的代码，只不过嵌进了方法体
        public void Connet(string Iptxt, int Port)//接收参数是目标ip地址和目标端口号。客户端无须关心本地端口号
        {
            //创建一个新的Socket对象
            // TcpClient 
            TcpClient clientSocket = new TcpClient();
            //clientSocket.Connect(tb_ServerIP.Text, Convert.ToInt32(tb_ServerPort.Text));
           // Socket client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IsConnet = true;//注意，此处是全局变量，将其设置为true
                            //将方法写进线程中
            Thread thread = new Thread(() =>
            {
                while (IsConnet)//循环
                {
                    try
                    {
                        clientSocket.Connect(Iptxt, Port);//尝试连接，失败则会跳去catch
                        if(clientSocket.Connected)
                        {
                            // 发送消息
                            string message = "4C4F4E0D";
                            byte[] data = HexStringToByteArray(message);
                            //NetworkStream stream = new NetworkStream();
                            //stream.Write();
                            NetworkStream stream = clientSocket.GetStream();

                            stream.Write(data, 0, data.Length);

                            // 接收服务器回复

                            byte[] responseData = new byte[256];
                            StringBuilder responseMessage = new StringBuilder();

                            int bytes = stream.Read(responseData, 0, responseData.Length);
                            responseMessage.Append(Encoding.ASCII.GetString(responseData, 0, bytes));

                            //Console.WriteLine("Received response from server: " + responseMessage);
                            tb_RecieveMsg.Items.Add("Received response from server: " + responseMessage);

                        }
                        IsConnet = false;//成功连接后修改bool值为false,这样下一步循环就不再执行。
                        break;//在此处加上break，成功就跳出循环，避免死循环
                    }
                    catch
                    {
                        clientSocket.Close();//先关闭
                        /*使用新的客户端资源覆盖，上一个已经废弃。如果继续使用以前的资源进行连接，
                        即使参数正确， 服务器全部打开也会无法连接*/
                        clientSocket = new TcpClient();
                        Thread.Sleep(1000);//等待1s再去重连
                    }
                }
                /*这里不一样就是放接收线程，在连接上后break出来，执行。
                因为需要带参数，所以要用到特别的ParameterizedThreadStart，
                然后开始线程。↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓*/
                Thread thread2 = new Thread(new ParameterizedThreadStart(ClientReceiveData));//接收线程方法
                thread2.IsBackground = true;//该值指示某个线程是否为后台线程。
                thread2.Start(clientSocket);//参数是用我们自建的Socket对象，就是上面的Socket client=new……

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
                    Connet("192.168.7.241", 9004);//重新尝试去连接
                    IsConnet = false;//注意，此处是全局变量，将其设置为false,防止循环
                    return;//让方法结束，终结当前接收服务端数据的异步线程
                }
                if (len <= 0)
                {
                    //如果小于0，证明无连接，服务端正常退出
                    ProxSocket.Shutdown(SocketShutdown.Both);//中止传输
                    ProxSocket.Close();//关闭
                    Connet("192.168.7.241", 9004);//重新尝试去连接
                    IsConnet = false;//注意，此处是全局变量，将其设置为false,防止循环
                    return;//让方法结束，终结当前接收服务端数据的异步线程
                }
                //这里做你想要对消息做的处理
                //string str = Encoding.Default.GetString(data, 0, len);//二进制数组转换成字符串……
            }
        }

        private void btn_Send_Click(object sender, EventArgs e)
        {
            Connet(tb_ServerIP.Text, Convert.ToInt32(tb_ServerPort.Text));
        }

        private void tb_ServerIP_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
