namespace TestClient
{
    partial class TestFrm
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
            this.tb_RecieveMsg = new System.Windows.Forms.ListBox();
            this.btn_Send = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.tb_SendMsg = new System.Windows.Forms.TextBox();
            this.btn_Link = new System.Windows.Forms.Button();
            this.tb_ServerPort = new System.Windows.Forms.TextBox();
            this.tb_ServerIP = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // tb_RecieveMsg
            // 
            this.tb_RecieveMsg.FormattingEnabled = true;
            this.tb_RecieveMsg.HorizontalExtent = 1;
            this.tb_RecieveMsg.ItemHeight = 12;
            this.tb_RecieveMsg.Location = new System.Drawing.Point(35, 187);
            this.tb_RecieveMsg.Margin = new System.Windows.Forms.Padding(2);
            this.tb_RecieveMsg.Name = "tb_RecieveMsg";
            this.tb_RecieveMsg.Size = new System.Drawing.Size(1044, 160);
            this.tb_RecieveMsg.TabIndex = 47;
            // 
            // btn_Send
            // 
            this.btn_Send.Location = new System.Drawing.Point(293, 362);
            this.btn_Send.Name = "btn_Send";
            this.btn_Send.Size = new System.Drawing.Size(75, 23);
            this.btn_Send.TabIndex = 46;
            this.btn_Send.Text = "发送信息";
            this.btn_Send.UseVisualStyleBackColor = true;
            this.btn_Send.Click += new System.EventHandler(this.btn_Send_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(33, 173);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 45;
            this.label4.Text = "收到的信息";
            // 
            // tb_SendMsg
            // 
            this.tb_SendMsg.Location = new System.Drawing.Point(35, 362);
            this.tb_SendMsg.Multiline = true;
            this.tb_SendMsg.Name = "tb_SendMsg";
            this.tb_SendMsg.Size = new System.Drawing.Size(237, 112);
            this.tb_SendMsg.TabIndex = 44;
            this.tb_SendMsg.Text = "LON";
            // 
            // btn_Link
            // 
            this.btn_Link.Location = new System.Drawing.Point(157, 122);
            this.btn_Link.Name = "btn_Link";
            this.btn_Link.Size = new System.Drawing.Size(75, 23);
            this.btn_Link.TabIndex = 43;
            this.btn_Link.Text = "连接服务器";
            this.btn_Link.UseVisualStyleBackColor = true;
            // 
            // tb_ServerPort
            // 
            this.tb_ServerPort.Location = new System.Drawing.Point(133, 69);
            this.tb_ServerPort.Name = "tb_ServerPort";
            this.tb_ServerPort.Size = new System.Drawing.Size(100, 22);
            this.tb_ServerPort.TabIndex = 41;
            this.tb_ServerPort.Text = "9004";
            // 
            // tb_ServerIP
            // 
            this.tb_ServerIP.Location = new System.Drawing.Point(133, 20);
            this.tb_ServerIP.Name = "tb_ServerIP";
            this.tb_ServerIP.Size = new System.Drawing.Size(100, 22);
            this.tb_ServerIP.TabIndex = 42;
            this.tb_ServerIP.Text = "192.168.7.240";
            this.tb_ServerIP.TextChanged += new System.EventHandler(this.tb_ServerIP_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(76, 72);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 40;
            this.label3.Text = "端口号";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(64, 29);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 39;
            this.label2.Text = "服务器地址";
            // 
            // TestFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1219, 667);
            this.Controls.Add(this.tb_RecieveMsg);
            this.Controls.Add(this.btn_Send);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tb_SendMsg);
            this.Controls.Add(this.btn_Link);
            this.Controls.Add(this.tb_ServerPort);
            this.Controls.Add(this.tb_ServerIP);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Name = "TestFrm";
            this.Text = "TestFrm";
            this.Load += new System.EventHandler(this.TestFrm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox tb_RecieveMsg;
        private System.Windows.Forms.Button btn_Send;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tb_SendMsg;
        private System.Windows.Forms.Button btn_Link;
        private System.Windows.Forms.TextBox tb_ServerPort;
        private System.Windows.Forms.TextBox tb_ServerIP;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
    }
}