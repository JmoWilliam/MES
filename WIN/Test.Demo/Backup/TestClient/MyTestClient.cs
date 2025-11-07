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
namespace TestClient
{
    public partial class MyTestClient : Form
    {
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
            ClientSocket.Close();

            this.btn_BreakLink.Enabled = false;
            this.btn_Link.Enabled = true;
        }

        private void btn_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Send_Click(object sender, EventArgs e)
        {
            MsgSend = Encoding.Unicode.GetBytes(this.tb_UserName.Text + "说：\n" + this.tb_SendMsg.Text + "\n");
            if (ClientSocket.Connected)
            {
                ClientSocket.Send(MsgSend);
                this.tb_SendMsg.Text = "";
            }
            else
            {
                MessageBox.Show("当前与服务器断开连接，无法发送信息！");
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

            this.tb_ServerIP.Text = "127.0.0.1";
            this.tb_ServerPort.Text = "6600";
            CheckForIllegalCrossThreadCalls = false;

        }
    }
}