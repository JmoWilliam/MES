namespace TestClient
{
    partial class SocketFrm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tb_ServerPort = new System.Windows.Forms.TextBox();
            this.tb_ServerIP = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_Link = new System.Windows.Forms.Button();
            this.tb_RecieveMsg = new System.Windows.Forms.ListBox();
            this.btn_Send = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.tb_SendMsg = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // tb_ServerPort
            // 
            this.tb_ServerPort.Enabled = false;
            this.tb_ServerPort.Location = new System.Drawing.Point(386, 23);
            this.tb_ServerPort.Margin = new System.Windows.Forms.Padding(4);
            this.tb_ServerPort.Name = "tb_ServerPort";
            this.tb_ServerPort.Size = new System.Drawing.Size(148, 29);
            this.tb_ServerPort.TabIndex = 29;
            this.tb_ServerPort.Text = "9004";
            // 
            // tb_ServerIP
            // 
            this.tb_ServerIP.Enabled = false;
            this.tb_ServerIP.Location = new System.Drawing.Point(121, 23);
            this.tb_ServerIP.Margin = new System.Windows.Forms.Padding(4);
            this.tb_ServerIP.Name = "tb_ServerIP";
            this.tb_ServerIP.Size = new System.Drawing.Size(148, 29);
            this.tb_ServerIP.TabIndex = 30;
            this.tb_ServerIP.Text = "192.168.26.201";
            this.tb_ServerIP.TextChanged += new System.EventHandler(this.tb_ServerIP_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(318, 33);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 18);
            this.label3.TabIndex = 28;
            this.label3.Text = "端口号";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 32);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 18);
            this.label2.TabIndex = 27;
            this.label2.Text = "服务器地址";
            // 
            // btn_Link
            // 
            this.btn_Link.Location = new System.Drawing.Point(578, 18);
            this.btn_Link.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Link.Name = "btn_Link";
            this.btn_Link.Size = new System.Drawing.Size(112, 34);
            this.btn_Link.TabIndex = 31;
            this.btn_Link.Text = "连接服务器";
            this.btn_Link.UseVisualStyleBackColor = true;
            this.btn_Link.Visible = false;
            this.btn_Link.Click += new System.EventHandler(this.btn_Link_Click);
            // 
            // tb_RecieveMsg
            // 
            this.tb_RecieveMsg.FormattingEnabled = true;
            this.tb_RecieveMsg.HorizontalScrollbar = true;
            this.tb_RecieveMsg.ItemHeight = 18;
            this.tb_RecieveMsg.Location = new System.Drawing.Point(24, 124);
            this.tb_RecieveMsg.Name = "tb_RecieveMsg";
            this.tb_RecieveMsg.Size = new System.Drawing.Size(904, 364);
            this.tb_RecieveMsg.TabIndex = 38;
            this.tb_RecieveMsg.SelectedIndexChanged += new System.EventHandler(this.tb_RecieveMsg_SelectedIndexChanged);
            // 
            // btn_Send
            // 
            this.btn_Send.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Send.Location = new System.Drawing.Point(388, 498);
            this.btn_Send.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Send.Name = "btn_Send";
            this.btn_Send.Size = new System.Drawing.Size(146, 68);
            this.btn_Send.TabIndex = 37;
            this.btn_Send.Text = "发送信息";
            this.btn_Send.UseVisualStyleBackColor = true;
            this.btn_Send.Click += new System.EventHandler(this.btn_Send_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 93);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 18);
            this.label4.TabIndex = 36;
            this.label4.Text = "收到的信息";
            // 
            // tb_SendMsg
            // 
            this.tb_SendMsg.Location = new System.Drawing.Point(24, 498);
            this.tb_SendMsg.Margin = new System.Windows.Forms.Padding(4);
            this.tb_SendMsg.Multiline = true;
            this.tb_SendMsg.Name = "tb_SendMsg";
            this.tb_SendMsg.Size = new System.Drawing.Size(354, 66);
            this.tb_SendMsg.TabIndex = 35;
            this.tb_SendMsg.Text = "LON";
            // 
            // SocketFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(944, 582);
            this.Controls.Add(this.tb_RecieveMsg);
            this.Controls.Add(this.btn_Send);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tb_SendMsg);
            this.Controls.Add(this.btn_Link);
            this.Controls.Add(this.tb_ServerPort);
            this.Controls.Add(this.tb_ServerIP);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Name = "SocketFrm";
            this.Text = "SocketFrm";
            this.Load += new System.EventHandler(this.SocketFrm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb_ServerPort;
        private System.Windows.Forms.TextBox tb_ServerIP;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_Link;
        private System.Windows.Forms.ListBox tb_RecieveMsg;
        private System.Windows.Forms.Button btn_Send;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tb_SendMsg;
    }
}