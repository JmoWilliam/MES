namespace SDKSimpleApp
{
    partial class Form1
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
            this.label5 = new System.Windows.Forms.Label();
            this.TxBarCode = new System.Windows.Forms.TextBox();
            this.labMFG = new System.Windows.Forms.Label();
            this.txMFG = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.TxUser = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtUserPassword = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(12, 180);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(116, 31);
            this.label5.TabIndex = 8;
            this.label5.Text = "模仁條碼:";
            // 
            // TxBarCode
            // 
            this.TxBarCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxBarCode.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxBarCode.Location = new System.Drawing.Point(134, 175);
            this.TxBarCode.MaxLength = 100;
            this.TxBarCode.Name = "TxBarCode";
            this.TxBarCode.Size = new System.Drawing.Size(549, 36);
            this.TxBarCode.TabIndex = 9;
            this.TxBarCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxBarCode_KeyDown);
            // 
            // labMFG
            // 
            this.labMFG.AutoSize = true;
            this.labMFG.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labMFG.Location = new System.Drawing.Point(12, 277);
            this.labMFG.Name = "labMFG";
            this.labMFG.Size = new System.Drawing.Size(116, 31);
            this.labMFG.TabIndex = 10;
            this.labMFG.Text = "製令條碼:";
            // 
            // txMFG
            // 
            this.txMFG.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txMFG.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txMFG.Location = new System.Drawing.Point(134, 272);
            this.txMFG.MaxLength = 100;
            this.txMFG.Name = "txMFG";
            this.txMFG.Size = new System.Drawing.Size(549, 36);
            this.txMFG.TabIndex = 11;
            this.txMFG.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txMFG_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(12, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 31);
            this.label1.TabIndex = 12;
            this.label1.Text = "員工工號:";
            // 
            // TxUser
            // 
            this.TxUser.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxUser.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TxUser.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxUser.Location = new System.Drawing.Point(134, 13);
            this.TxUser.MaxLength = 100;
            this.TxUser.Name = "TxUser";
            this.TxUser.Size = new System.Drawing.Size(549, 36);
            this.TxUser.TabIndex = 13;
            this.TxUser.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxUser_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(12, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 31);
            this.label2.TabIndex = 14;
            this.label2.Text = "員工密碼:";
            // 
            // txtUserPassword
            // 
            this.txtUserPassword.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtUserPassword.Location = new System.Drawing.Point(134, 90);
            this.txtUserPassword.Name = "txtUserPassword";
            this.txtUserPassword.Size = new System.Drawing.Size(546, 39);
            this.txtUserPassword.TabIndex = 15;
            this.txtUserPassword.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUserPassword_Keydown);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(692, 381);
            this.Controls.Add(this.txtUserPassword);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TxUser);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txMFG);
            this.Controls.Add(this.labMFG);
            this.Controls.Add(this.TxBarCode);
            this.Controls.Add(this.label5);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "舜宇出貨標籤";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox TxBarCode;
        private System.Windows.Forms.Label labMFG;
        private System.Windows.Forms.TextBox txMFG;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TxUser;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtUserPassword;
    }
}

