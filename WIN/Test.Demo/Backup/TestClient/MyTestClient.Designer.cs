namespace TestClient
{
    partial class MyTestClient
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSendGameMsg = new System.Windows.Forms.Button();
            this.btnBreakGame = new System.Windows.Forms.Button();
            this.btnJoinGame = new System.Windows.Forms.Button();
            this.btn_Send = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.tb_SendMsg = new System.Windows.Forms.TextBox();
            this.tb_RecieveMsg = new System.Windows.Forms.TextBox();
            this.tb_ServerPort = new System.Windows.Forms.TextBox();
            this.tb_ServerIP = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tb_UserName = new System.Windows.Forms.TextBox();
            this.btn_Exit = new System.Windows.Forms.Button();
            this.btn_BreakLink = new System.Windows.Forms.Button();
            this.btn_Link = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSendGameMsg);
            this.groupBox1.Controls.Add(this.btnBreakGame);
            this.groupBox1.Controls.Add(this.btnJoinGame);
            this.groupBox1.Location = new System.Drawing.Point(14, 388);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(417, 189);
            this.groupBox1.TabIndex = 32;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "测试";
            // 
            // btnSendGameMsg
            // 
            this.btnSendGameMsg.Location = new System.Drawing.Point(6, 78);
            this.btnSendGameMsg.Name = "btnSendGameMsg";
            this.btnSendGameMsg.Size = new System.Drawing.Size(102, 23);
            this.btnSendGameMsg.TabIndex = 2;
            this.btnSendGameMsg.Text = "发送游戏数据包";
            this.btnSendGameMsg.UseVisualStyleBackColor = true;
            // 
            // btnBreakGame
            // 
            this.btnBreakGame.Location = new System.Drawing.Point(6, 49);
            this.btnBreakGame.Name = "btnBreakGame";
            this.btnBreakGame.Size = new System.Drawing.Size(75, 23);
            this.btnBreakGame.TabIndex = 1;
            this.btnBreakGame.Text = "断开游戏";
            this.btnBreakGame.UseVisualStyleBackColor = true;
            // 
            // btnJoinGame
            // 
            this.btnJoinGame.Location = new System.Drawing.Point(6, 20);
            this.btnJoinGame.Name = "btnJoinGame";
            this.btnJoinGame.Size = new System.Drawing.Size(75, 23);
            this.btnJoinGame.TabIndex = 0;
            this.btnJoinGame.Text = "加入游戏";
            this.btnJoinGame.UseVisualStyleBackColor = true;
            // 
            // btn_Send
            // 
            this.btn_Send.Location = new System.Drawing.Point(14, 358);
            this.btn_Send.Name = "btn_Send";
            this.btn_Send.Size = new System.Drawing.Size(75, 23);
            this.btn_Send.TabIndex = 31;
            this.btn_Send.Text = "发送信息";
            this.btn_Send.UseVisualStyleBackColor = true;
            this.btn_Send.Click += new System.EventHandler(this.btn_Send_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 196);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 29;
            this.label5.Text = "发送的信息";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 7);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 30;
            this.label4.Text = "收到的信息";
            // 
            // tb_SendMsg
            // 
            this.tb_SendMsg.Location = new System.Drawing.Point(12, 222);
            this.tb_SendMsg.Multiline = true;
            this.tb_SendMsg.Name = "tb_SendMsg";
            this.tb_SendMsg.Size = new System.Drawing.Size(237, 112);
            this.tb_SendMsg.TabIndex = 28;
            // 
            // tb_RecieveMsg
            // 
            this.tb_RecieveMsg.Location = new System.Drawing.Point(12, 33);
            this.tb_RecieveMsg.Multiline = true;
            this.tb_RecieveMsg.Name = "tb_RecieveMsg";
            this.tb_RecieveMsg.Size = new System.Drawing.Size(237, 144);
            this.tb_RecieveMsg.TabIndex = 27;
            // 
            // tb_ServerPort
            // 
            this.tb_ServerPort.Location = new System.Drawing.Point(331, 76);
            this.tb_ServerPort.Name = "tb_ServerPort";
            this.tb_ServerPort.Size = new System.Drawing.Size(100, 21);
            this.tb_ServerPort.TabIndex = 25;
            // 
            // tb_ServerIP
            // 
            this.tb_ServerIP.Location = new System.Drawing.Point(331, 27);
            this.tb_ServerIP.Name = "tb_ServerIP";
            this.tb_ServerIP.Size = new System.Drawing.Size(100, 21);
            this.tb_ServerIP.TabIndex = 26;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(274, 79);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 24;
            this.label3.Text = "端口号";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(262, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 23;
            this.label2.Text = "服务器地址";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(274, 121);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 22;
            this.label1.Text = "用户昵称";
            // 
            // tb_UserName
            // 
            this.tb_UserName.Location = new System.Drawing.Point(331, 113);
            this.tb_UserName.Name = "tb_UserName";
            this.tb_UserName.Size = new System.Drawing.Size(100, 21);
            this.tb_UserName.TabIndex = 21;
            // 
            // btn_Exit
            // 
            this.btn_Exit.Location = new System.Drawing.Point(356, 358);
            this.btn_Exit.Name = "btn_Exit";
            this.btn_Exit.Size = new System.Drawing.Size(75, 23);
            this.btn_Exit.TabIndex = 20;
            this.btn_Exit.Text = "退出";
            this.btn_Exit.UseVisualStyleBackColor = true;
            this.btn_Exit.Click += new System.EventHandler(this.btn_Exit_Click);
            // 
            // btn_BreakLink
            // 
            this.btn_BreakLink.Enabled = false;
            this.btn_BreakLink.Location = new System.Drawing.Point(356, 329);
            this.btn_BreakLink.Name = "btn_BreakLink";
            this.btn_BreakLink.Size = new System.Drawing.Size(75, 23);
            this.btn_BreakLink.TabIndex = 19;
            this.btn_BreakLink.Text = "断开连接";
            this.btn_BreakLink.UseVisualStyleBackColor = true;
            this.btn_BreakLink.Click += new System.EventHandler(this.btn_BreakLink_Click);
            // 
            // btn_Link
            // 
            this.btn_Link.Location = new System.Drawing.Point(356, 300);
            this.btn_Link.Name = "btn_Link";
            this.btn_Link.Size = new System.Drawing.Size(75, 23);
            this.btn_Link.TabIndex = 18;
            this.btn_Link.Text = "连接服务器";
            this.btn_Link.UseVisualStyleBackColor = true;
            this.btn_Link.Click += new System.EventHandler(this.btn_Link_Click);
            // 
            // MyTestClient
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(454, 593);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btn_Send);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tb_SendMsg);
            this.Controls.Add(this.tb_RecieveMsg);
            this.Controls.Add(this.tb_ServerPort);
            this.Controls.Add(this.tb_ServerIP);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tb_UserName);
            this.Controls.Add(this.btn_Exit);
            this.Controls.Add(this.btn_BreakLink);
            this.Controls.Add(this.btn_Link);
            this.Name = "MyTestClient";
            this.Text = "MyTestClient";
            this.Load += new System.EventHandler(this.MyTestClient_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnSendGameMsg;
        private System.Windows.Forms.Button btnBreakGame;
        private System.Windows.Forms.Button btnJoinGame;
        private System.Windows.Forms.Button btn_Send;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tb_SendMsg;
        private System.Windows.Forms.TextBox tb_RecieveMsg;
        private System.Windows.Forms.TextBox tb_ServerPort;
        private System.Windows.Forms.TextBox tb_ServerIP;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_UserName;
        private System.Windows.Forms.Button btn_Exit;
        private System.Windows.Forms.Button btn_BreakLink;
        private System.Windows.Forms.Button btn_Link;
    }
}