namespace TestClient
{
    partial class TClient
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnLink = new System.Windows.Forms.Button();
            this.btnBreakLink = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.ServerPort = new System.Windows.Forms.TextBox();
            this.ServerIP = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.UserName = new System.Windows.Forms.TextBox();
            this.SendMsg = new System.Windows.Forms.TextBox();
            this.RecieveMsg = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnSend = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnJoinGame = new System.Windows.Forms.Button();
            this.btnBreakGame = new System.Windows.Forms.Button();
            this.btnSendGameMsg = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnLink
            // 
            this.btnLink.Location = new System.Drawing.Point(356, 306);
            this.btnLink.Name = "btnLink";
            this.btnLink.Size = new System.Drawing.Size(75, 23);
            this.btnLink.TabIndex = 0;
            this.btnLink.Text = "连接服务器";
            this.btnLink.UseVisualStyleBackColor = true;
            this.btnLink.Click += new System.EventHandler(this.btnLink_Click);
            // 
            // btnBreakLink
            // 
            this.btnBreakLink.Enabled = false;
            this.btnBreakLink.Location = new System.Drawing.Point(356, 335);
            this.btnBreakLink.Name = "btnBreakLink";
            this.btnBreakLink.Size = new System.Drawing.Size(75, 23);
            this.btnBreakLink.TabIndex = 1;
            this.btnBreakLink.Text = "断开连接";
            this.btnBreakLink.UseVisualStyleBackColor = true;
            this.btnBreakLink.Click += new System.EventHandler(this.btnBreakLink_Click);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(356, 364);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "退出";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // ServerPort
            // 
            this.ServerPort.Location = new System.Drawing.Point(331, 82);
            this.ServerPort.Name = "ServerPort";
            this.ServerPort.Size = new System.Drawing.Size(100, 21);
            this.ServerPort.TabIndex = 11;
            // 
            // ServerIP
            // 
            this.ServerIP.Location = new System.Drawing.Point(331, 33);
            this.ServerIP.Name = "ServerIP";
            this.ServerIP.Size = new System.Drawing.Size(100, 21);
            this.ServerIP.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(274, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 10;
            this.label3.Text = "端口号";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(262, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "服务器地址";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(274, 127);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "用户昵称";
            // 
            // UserName
            // 
            this.UserName.Location = new System.Drawing.Point(331, 119);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(100, 21);
            this.UserName.TabIndex = 7;
            // 
            // SendMsg
            // 
            this.SendMsg.Location = new System.Drawing.Point(12, 228);
            this.SendMsg.Multiline = true;
            this.SendMsg.Name = "SendMsg";
            this.SendMsg.Size = new System.Drawing.Size(237, 112);
            this.SendMsg.TabIndex = 14;
            // 
            // RecieveMsg
            // 
            this.RecieveMsg.Location = new System.Drawing.Point(12, 39);
            this.RecieveMsg.Multiline = true;
            this.RecieveMsg.Name = "RecieveMsg";
            this.RecieveMsg.Size = new System.Drawing.Size(237, 144);
            this.RecieveMsg.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 13);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 15;
            this.label4.Text = "收到的信息";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 202);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 15;
            this.label5.Text = "发送的信息";
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(14, 364);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 23);
            this.btnSend.TabIndex = 16;
            this.btnSend.Text = "发送信息";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSendGameMsg);
            this.groupBox1.Controls.Add(this.btnBreakGame);
            this.groupBox1.Controls.Add(this.btnJoinGame);
            this.groupBox1.Location = new System.Drawing.Point(14, 394);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(417, 189);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "测试";
            // 
            // btnJoinGame
            // 
            this.btnJoinGame.Location = new System.Drawing.Point(6, 20);
            this.btnJoinGame.Name = "btnJoinGame";
            this.btnJoinGame.Size = new System.Drawing.Size(75, 23);
            this.btnJoinGame.TabIndex = 0;
            this.btnJoinGame.Text = "加入游戏";
            this.btnJoinGame.UseVisualStyleBackColor = true;
            this.btnJoinGame.Click += new System.EventHandler(this.btnJoinGame_Click);
            // 
            // btnBreakGame
            // 
            this.btnBreakGame.Location = new System.Drawing.Point(6, 49);
            this.btnBreakGame.Name = "btnBreakGame";
            this.btnBreakGame.Size = new System.Drawing.Size(75, 23);
            this.btnBreakGame.TabIndex = 1;
            this.btnBreakGame.Text = "断开游戏";
            this.btnBreakGame.UseVisualStyleBackColor = true;
            this.btnBreakGame.Click += new System.EventHandler(this.btnBreakGame_Click);
            // 
            // btnSendGameMsg
            // 
            this.btnSendGameMsg.Location = new System.Drawing.Point(6, 78);
            this.btnSendGameMsg.Name = "btnSendGameMsg";
            this.btnSendGameMsg.Size = new System.Drawing.Size(102, 23);
            this.btnSendGameMsg.TabIndex = 2;
            this.btnSendGameMsg.Text = "发送游戏数据包";
            this.btnSendGameMsg.UseVisualStyleBackColor = true;
            this.btnSendGameMsg.Click += new System.EventHandler(this.btnSendGameMsg_Click);
            // 
            // TClient
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(443, 595);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.SendMsg);
            this.Controls.Add(this.RecieveMsg);
            this.Controls.Add(this.ServerPort);
            this.Controls.Add(this.ServerIP);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.UserName);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnBreakLink);
            this.Controls.Add(this.btnLink);
            this.Name = "TClient";
            this.Text = "测试客户端";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.TClient_FormClosing);
            this.Load += new System.EventHandler(this.TClient_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLink;
        private System.Windows.Forms.Button btnBreakLink;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.TextBox ServerPort;
        private System.Windows.Forms.TextBox ServerIP;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.TextBox SendMsg;
        private System.Windows.Forms.TextBox RecieveMsg;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnBreakGame;
        private System.Windows.Forms.Button btnJoinGame;
        private System.Windows.Forms.Button btnSendGameMsg;
    }
}

