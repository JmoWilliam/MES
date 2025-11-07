using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using Keyence.AutoID.SDK;
using System.Xml;
using System.Threading;
using Keyence.AR.Communication;

namespace TestClient
{
    public partial class MyTestClient : Form
    {
        private ReaderAccessor m_reader = new ReaderAccessor();
        private object _lockEventTcpIp = new object();
        private NetworkRequestClient _client = new NetworkRequestClient();
        public ErrorCode LastErrorInfo { get; private set; }

        private IPEndPoint ServerInfo;//存放服务器的IP和端口信息
        private Socket ClientSocket;    //客户端SOCKET
        private Byte[] MsgBuffer;       //存放消息数据缓冲
        private Byte[] MsgSend;         //发送消息数据
        public MyTestClient()
        {
            InitializeComponent();
        }
  
        private void btn_Link_Click(object sender, EventArgs e)
        {
            ClientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            ServerInfo = new IPEndPoint(IPAddress.Parse(this.tb_ServerIP.Text), Convert.ToInt32(this.tb_ServerPort.Text));
            //try
            //{
            ClientSocket.Connect(ServerInfo);
            ClientSocket.Send(Encoding.Unicode.GetBytes("用户： " + this.tb_UserName.Text + " 进入系统！\n"));
            ClientSocket.BeginReceive(MsgBuffer, 0, MsgBuffer.Length, 0, new AsyncCallback(ReceiveCallBack), null);



            this.btn_BreakLink.Enabled = true;
                this.btn_Link.Enabled = false;
                this.btn_Send.Enabled = true;

                MessageBox.Show("登录服务器成功！\n");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
        //回发数据
        private void ReceiveCallBack(IAsyncResult AR)
        {
            try
            {
                int REnd = ClientSocket.EndReceive(AR);
                this.tb_SendMsg.AppendText(Encoding.Unicode.GetString(MsgBuffer, 0, REnd)+"\r\n");
                ClientSocket.BeginReceive(MsgBuffer, 0, MsgBuffer.Length, 0, new AsyncCallback(ReceiveCallBack), null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btn_BreakLink_Click(object sender, EventArgs e)
        {
           // ClientSocket.Disconnect();
            ClientSocket.Close();

            this.btn_BreakLink.Enabled = false;
            this.btn_Link.Enabled = true;
        }

        private void btn_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        #region C# ASCII转字符及字符转ASCII
        public static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();//转为
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);//转为字符
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }

        }

        public static int Asc(string character)
        {
            if (character.Length == 1)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int)asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                throw new Exception("Character is not valid.");
            }

        }
        #endregion

        private delegate void delegateUserControl(string str);
        private void ReceivedDataWrite(string receivedData)
        {
            tb_RecieveMsg.Items.Add(("[" + tb_ServerIP.Text + "][" + DateTime.Now + "]" + receivedData));
        }

        private ErrorCode convertErrorInfo(CommunicationResult res)
        {
            if (res.Equals(CommunicationResult.AddressUsed))
            {
                return ErrorCode.IpAddressUsed;
            }

            if (res.Equals(CommunicationResult.AlreadyOpen))
            {
                return ErrorCode.AlreadyOpen;
            }

            if (res.Equals(CommunicationResult.Closed))
            {
                return ErrorCode.Closed;
            }

            if (res.Equals(CommunicationResult.OpenFailed))
            {
                return ErrorCode.OpenFailed;
            }

            if (res.Equals(CommunicationResult.HeadFailed))
            {
                return ErrorCode.HeadFailed;
            }

            if (res.Equals(CommunicationResult.SendFailed))
            {
                return ErrorCode.SendFailed;
            }

            if (res.Equals(CommunicationResult.Success))
            {
                return ErrorCode.None;
            }

            if (res.Equals(CommunicationResult.Timeout))
            {
                return ErrorCode.Timeout;
            }

            if ((int)res == 10048)
            {
                return ErrorCode.SocketAddressUsed;
            }

            if ((int)res == 10049)
            {
                return ErrorCode.SocketAddressDisabled;
            }

            if ((int)res == 10051)
            {
                return ErrorCode.SocketConnectionUnreach;
            }

            if ((int)res == 10054)
            {
                return ErrorCode.SocketConnectionReset;
            }

            if ((int)res == 10056)
            {
                return ErrorCode.SocketAlreadyConnected;
            }

            if ((int)res == 10060)
            {
                return ErrorCode.SocketConnectionTimeout;
            }

            if ((int)res == 10061)
            {
                return ErrorCode.SocketConnectionRefused;
            }

            if ((int)res == 421)
            {
                return ErrorCode.FtpServiceUnavailable;
            }

            if ((int)res == 425)
            {
                return ErrorCode.FtpCannotOpenDataConnection;
            }

            if ((int)res == 426)
            {
                return ErrorCode.FtpDataConnectionDisconnected;
            }

            if ((int)res == 450)
            {
                return ErrorCode.FtpFileBusy;
            }

            if ((int)res == 451)
            {
                return ErrorCode.FtpActionAborted;
            }

            if ((int)res == 452)
            {
                return ErrorCode.FtpDiskFull;
            }

            if ((int)res == 500)
            {
                return ErrorCode.FtpCommandUnrecognized;
            }

            if ((int)res == 501)
            {
                return ErrorCode.FtpInvalidArgument;
            }

            if ((int)res == 502)
            {
                return ErrorCode.FtpCommandUnimplemented;
            }

            if ((int)res == 503)
            {
                return ErrorCode.FtpCommandBadSequence;
            }

            if ((int)res == 504)
            {
                return ErrorCode.FtpArgumentsUnimplemented;
            }

            if ((int)res == 530)
            {
                return ErrorCode.FtpNotLoggedIn;
            }

            if ((int)res == 550)
            {
                return ErrorCode.FtpActionFailed;
            }

            if ((int)res == 552)
            {
                return ErrorCode.FtpExceededDisk;
            }

            if ((int)res == 553)
            {
                return ErrorCode.FtpFileActionFailed;
            }

            return ErrorCode.UnexpectedError;
        }

      
        private void btn_Send_Click(object sender, EventArgs e)
        {
           // MsgSend = Encoding.Unicode.GetBytes(this.tb_UserName.Text + "说：\n" + this.tb_SendMsg.Text + "\n");
            if (ClientSocket.Connected)
            {


                MsgSend =  System.Text.Encoding.ASCII.GetBytes("LON"+ "\r");
                int recv;
                int receiveLength = 0;
                int index = 0;
                ClientSocket.Send(MsgSend);
                System.Threading.Thread.Sleep(1000);
                ReceiveMessage(MsgSend);
                //while (ClientSocket.Available>0)
                //{
                //    //参数 数据缓存区  起始位置  数据长度  值的按位组合
                //    receiveLength += ClientSocket.Receive(MsgSend, index, ClientSocket.ReceiveBufferSize, SocketFlags.None);
                //    index += receiveLength;

                //    //return Encoding.GetEncoding("GB18030").GetString(result, 0, index);




                //    this.tb_RecieveMsg.Invoke((EventHandler)
                //        /*表示一个委托，该委托可执行托管代码中声明为 void 且不接受任何参数的任何方法。 在对控件的 Invoke 	方法进行调用时或需要一个简单委托又不想自己定义时可以使用该委托。*/
                //        delegate
                //        {

                //            this.tb_RecieveMsg.Text+= "";

                //        }
                //        );
                //    //BeginInvoke(new delegateUserControl(ReceivedDataWrite), System.Text.Encoding.UTF8.GetString(MsgSend));
                //}
            }
            else
            {
                MessageBox.Show("当前与服务器断开连接，无法发送信息！");
            }
        }

        private  void ReceiveMessage(byte[] MsgSend)
        {
            while (true)
            {
                try
                {
                    //通过clientSocket接收数据
                    int receiveNumber = ClientSocket.Receive(MsgSend);
                    string strContent = Encoding.ASCII.GetString(MsgSend, 0, receiveNumber);
                    tb_RecieveMsg.Items.Add(strContent);
                    //Console.WriteLine("接收服务端{0}消息{1}", ClientSocket.RemoteEndPoint.ToString(), Encoding.ASCII.GetString(MsgSend, 0, receiveNumber));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    ClientSocket.Shutdown(SocketShutdown.Both);
                    ClientSocket.Close();
                    break;
                }
            }
        }

        private void MyTestClient_Load(object sender, EventArgs e)
        {
            this.btn_Send.Enabled = false;
            this.btn_BreakLink.Enabled = false;


            ClientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            MsgBuffer = new Byte[65535];
            MsgSend = new Byte[65535];



            Random TRand = new Random();
            this.tb_UserName.Text = TRand.Next(10000).ToString();

            this.tb_ServerIP.Text = "169.254.0.1";
            this.tb_ServerPort.Text = "9004";
            CheckForIllegalCrossThreadCalls = false;

        }


        private void btnJoinGame_Click(object sender, EventArgs e)
        {

        }

        private void btnBreakGame_Click(object sender, EventArgs e)
        {

        }

        private void btnSendGameMsg_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            m_reader.IpAddress = "169.254.0.1";//comboBox1.SelectedItem.ToString();
                                               //Connect TCP/IP.
            m_reader.DataPort = 9004;
            ReceivedDataWrite(m_reader.ExecCommand("LON"));
            
            m_reader.Connect((data) =>
            {
                //Define received data actions here.Defined actions work asynchronously.
                //"ReceivedDataWrite" works when reading data was received.
                BeginInvoke(new delegateUserControl(ReceivedDataWrite), Encoding.ASCII.GetString(data));
            });
        }
    }
}