using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Net;
using System.Threading;
using System.Net.Sockets;
using PubTeacherClass;
using Keyence.AutoID.SDK;

namespace TestClient
{
    public partial class TClient : Form
    {
        public TClient()
        {
            InitializeComponent();
        }

        #region 定义变量

        private IPEndPoint ServerInfo;//存放服务器的IP和端口信息
        private Socket ClientSocket;    //客户端SOCKET
        private Byte[] MsgBuffer;       //存放消息数据缓冲
        private Byte[] MsgSend;         //发送消息数据

        private IPEndPoint ServerGameInfo;  //存放服务器游戏的端口信息
        private Socket ClientGameSocket;    //客户端游戏所用的SOCKET
        private Byte[] MsgGameBuffer;       //存放游戏消息数据缓冲
        private Byte[] MsgGameSend;         //发送游戏消息数据


        #endregion

        #region 事件响应

        //初始化加载
        private void TClient_Load(object sender, EventArgs e)
        {
            this.btnSend.Enabled = false;
            this.btnBreakLink.Enabled = false;


            ClientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            MsgBuffer = new Byte[65535];
            MsgSend = new Byte[65535];


            ClientGameSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            MsgGameBuffer = new Byte[65535];
            MsgGameSend = new Byte[65535];

            
            Random TRand = new Random();
            this.UserName.Text = TRand.Next(10000).ToString();

            this.ServerIP.Text = "127.0.0.1";
            this.ServerPort.Text = "6600";
            CheckForIllegalCrossThreadCalls = false;

       

        }

        //退出
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //关闭系统
        private void TClient_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("您确定要退出吗？  ", "温馨提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {

                Application.ExitThread();
                System.Diagnostics.Process[] LocalPro = System.Diagnostics.Process.GetProcessesByName("TestClient");
                if (LocalPro.Length > 0)
                {
                    foreach (System.Diagnostics.Process a in LocalPro)
                    {
                        a.Kill();
                    }
                }

            }
            else
            {
                e.Cancel = true;
            }
        }


        //连接服务器
        private void btnLink_Click(object sender, EventArgs e)
        {
            ClientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            ServerInfo = new IPEndPoint(IPAddress.Parse(this.ServerIP.Text), Convert.ToInt32(this.ServerPort.Text));
            try
            {
                ClientSocket.Connect(ServerInfo);
                ClientSocket.Send(Encoding.Unicode.GetBytes("用户： " + this.UserName.Text + " 进入系统！\n"));
                ClientSocket.BeginReceive(MsgBuffer, 0, MsgBuffer.Length, 0, new AsyncCallback(ReceiveCallBack), null);


                this.btnBreakLink.Enabled = true;
                this.btnLink.Enabled = false;
                this.btnSend.Enabled = true;

                MessageBox.Show("登录服务器成功！\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //断开服务器
        private void btnBreakLink_Click(object sender, EventArgs e)
        {
            //////////////if (ClientSocket.Connected)
            //////////////{
            //////////////    //ClientSocket.Send(Encoding.Unicode.GetBytes(this.UserName.Text + "离开了房间！\n"));
            //////////////    ClientSocket.Shutdown(SocketShutdown.Both);
            //////////////    ClientSocket.Disconnect(true);
            //////////////}
            ClientSocket.Close();

            this.btnBreakLink.Enabled = false;
            this.btnLink.Enabled = true;

            //MessageBox.Show("已经断开与服务器的连接...");
        }

        //发送普通消息
        private void btnSend_Click(object sender, EventArgs e)
        {
            MsgSend = Encoding.Unicode.GetBytes(this.UserName.Text + "说：\n" + this.SendMsg.Text + "\n");
            if (ClientSocket.Connected)
            {
                ClientSocket.Send(MsgSend);
                this.SendMsg.Text = "";
            }
            else
            {
                MessageBox.Show("当前与服务器断开连接，无法发送信息！");
            }
        }

        //加入游戏
        private void btnJoinGame_Click(object sender, EventArgs e)
        {
            ClientGameSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            ServerGameInfo = new IPEndPoint(IPAddress.Parse(this.ServerIP.Text), 11000);

            try
            {
                ClientGameSocket.Connect(ServerGameInfo);
                //ClientGameSocket.Send(Encoding.Unicode.GetBytes("用户： " + this.UserName.Text + " 加入游戏！\n"));
                ClientGameSocket.BeginReceive(MsgGameBuffer, 0, MsgGameBuffer.Length, 0, new AsyncCallback(ReceiveGameCallBack), null);

                this.btnBreakGame.Enabled = true;
                this.btnJoinGame.Enabled = false;

                MessageBox.Show("加入游戏成功！\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        //断开游戏
        private void btnBreakGame_Click(object sender, EventArgs e)
        {
            ClientGameSocket.Close();
            this.btnBreakGame.Enabled = false;
            this.btnJoinGame.Enabled = true;
        }

        #endregion

        #region 功能函数

        //回发数据
        private void ReceiveCallBack(IAsyncResult AR)
        {
            try
            {
                int REnd = ClientSocket.EndReceive(AR);
                this.RecieveMsg.AppendText(Encoding.Unicode.GetString(MsgBuffer, 0, REnd));
                ClientSocket.BeginReceive(MsgBuffer, 0, MsgBuffer.Length, 0, new AsyncCallback(ReceiveCallBack), null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //回发游戏数据
        private void ReceiveGameCallBack(IAsyncResult AR)
        {

            try
            {
                int REnd = ClientGameSocket.EndReceive(AR);
                this.RecieveMsg.AppendText(Encoding.Unicode.GetString(MsgGameBuffer, 0, REnd));

                Teacher t = (Teacher)Serializer.DeserializeObject(Convert.ToBase64String(MsgGameBuffer, 0, REnd));
                MessageBox.Show(t.ID + t.Name);    

                ClientGameSocket.BeginReceive(MsgGameBuffer, 0, MsgGameBuffer.Length, 0, new AsyncCallback(ReceiveGameCallBack), null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        #endregion

        #region 测试区

        //测试发送游戏数据
        private void btnSendGameMsg_Click(object sender, EventArgs e)
        {
            Teacher teacher = new Teacher("T100", "林林");



            MsgGameSend = Serializer.Serialize(teacher);
            if (ClientGameSocket.Connected)
            {
                ClientGameSocket.Send(MsgGameSend);
                //this.SendMsg.Text = "";
            }
            else
            {
                MessageBox.Show("当前与服务器断开连接，无法发送信息！");
            }
        }

        #endregion

        private void SendMsg_TextChanged(object sender, EventArgs e)
        {

        }
    }
}