namespace PrintLabelSystem
{
    partial class ToolLabel
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
            this.txtNotPrintLabelNum = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPartNotPrintLabelNum = new System.Windows.Forms.TextBox();
            this.labInputNun = new System.Windows.Forms.Label();
            this.txtKnifeModel_ID = new System.Windows.Forms.TextBox();
            this.labKnifeModelID = new System.Windows.Forms.Label();
            this.btnPrintLotLabel = new System.Windows.Forms.Button();
            this.txtToolNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnPrintLabel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtBarcodeType = new System.Windows.Forms.TextBox();
            this.txttxtBarcodeTypePcs = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(70, 151);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(130, 35);
            this.label5.TabIndex = 24;
            this.label5.Text = "條碼類型:";
            // 
            // txtNotPrintLabelNum
            // 
            this.txtNotPrintLabelNum.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtNotPrintLabelNum.Location = new System.Drawing.Point(236, 85);
            this.txtNotPrintLabelNum.Name = "txtNotPrintLabelNum";
            this.txtNotPrintLabelNum.Size = new System.Drawing.Size(427, 43);
            this.txtNotPrintLabelNum.TabIndex = 23;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(16, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(184, 35);
            this.label4.TabIndex = 22;
            this.label4.Text = "未列印的數量:";
            // 
            // txtPartNotPrintLabelNum
            // 
            this.txtPartNotPrintLabelNum.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtPartNotPrintLabelNum.Location = new System.Drawing.Point(236, 212);
            this.txtPartNotPrintLabelNum.Name = "txtPartNotPrintLabelNum";
            this.txtPartNotPrintLabelNum.Size = new System.Drawing.Size(425, 43);
            this.txtPartNotPrintLabelNum.TabIndex = 19;
            // 
            // labInputNun
            // 
            this.labInputNun.AutoSize = true;
            this.labInputNun.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labInputNun.Location = new System.Drawing.Point(70, 220);
            this.labInputNun.Name = "labInputNun";
            this.labInputNun.Size = new System.Drawing.Size(130, 35);
            this.labInputNun.TabIndex = 18;
            this.labInputNun.Text = "輸入數量:";
            // 
            // txtKnifeModel_ID
            // 
            this.txtKnifeModel_ID.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtKnifeModel_ID.Location = new System.Drawing.Point(236, 21);
            this.txtKnifeModel_ID.Name = "txtKnifeModel_ID";
            this.txtKnifeModel_ID.Size = new System.Drawing.Size(425, 43);
            this.txtKnifeModel_ID.TabIndex = 16;
            this.txtKnifeModel_ID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtKnifeModel_ID_KeyDown);
            // 
            // labKnifeModelID
            // 
            this.labKnifeModelID.AutoSize = true;
            this.labKnifeModelID.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labKnifeModelID.Location = new System.Drawing.Point(70, 27);
            this.labKnifeModelID.Name = "labKnifeModelID";
            this.labKnifeModelID.Size = new System.Drawing.Size(130, 35);
            this.labKnifeModelID.TabIndex = 15;
            this.labKnifeModelID.Text = "型號編號:";
            // 
            // btnPrintLotLabel
            // 
            this.btnPrintLotLabel.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnPrintLotLabel.Location = new System.Drawing.Point(267, 275);
            this.btnPrintLotLabel.Name = "btnPrintLotLabel";
            this.btnPrintLotLabel.Size = new System.Drawing.Size(112, 41);
            this.btnPrintLotLabel.TabIndex = 27;
            this.btnPrintLotLabel.Text = "批量列印";
            this.btnPrintLotLabel.UseVisualStyleBackColor = true;
            this.btnPrintLotLabel.Click += new System.EventHandler(this.btnPrintLotLabel_Click);
            // 
            // txtToolNo
            // 
            this.txtToolNo.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtToolNo.Location = new System.Drawing.Point(236, 21);
            this.txtToolNo.Name = "txtToolNo";
            this.txtToolNo.Size = new System.Drawing.Size(425, 43);
            this.txtToolNo.TabIndex = 29;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(89, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(130, 35);
            this.label2.TabIndex = 28;
            this.label2.Text = "工具條碼:";
            // 
            // btnPrintLabel
            // 
            this.btnPrintLabel.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnPrintLabel.Location = new System.Drawing.Point(245, 146);
            this.btnPrintLabel.Name = "btnPrintLabel";
            this.btnPrintLabel.Size = new System.Drawing.Size(112, 41);
            this.btnPrintLabel.TabIndex = 30;
            this.btnPrintLabel.Text = "單支列印";
            this.btnPrintLabel.UseVisualStyleBackColor = true;
            this.btnPrintLabel.Click += new System.EventHandler(this.btnPrintLabel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.groupBox1.Controls.Add(this.txtKnifeModel_ID);
            this.groupBox1.Controls.Add(this.txtNotPrintLabelNum);
            this.groupBox1.Controls.Add(this.txtBarcodeType);
            this.groupBox1.Controls.Add(this.labKnifeModelID);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txtPartNotPrintLabelNum);
            this.groupBox1.Controls.Add(this.btnPrintLotLabel);
            this.groupBox1.Controls.Add(this.labInputNun);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(683, 339);
            this.groupBox1.TabIndex = 31;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.groupBox2.Controls.Add(this.txttxtBarcodeTypePcs);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txtToolNo);
            this.groupBox2.Controls.Add(this.btnPrintLabel);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(12, 386);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(683, 213);
            this.groupBox2.TabIndex = 32;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "groupBox2";
            // 
            // txtBarcodeType
            // 
            this.txtBarcodeType.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBarcodeType.Location = new System.Drawing.Point(236, 143);
            this.txtBarcodeType.Name = "txtBarcodeType";
            this.txtBarcodeType.Size = new System.Drawing.Size(427, 43);
            this.txtBarcodeType.TabIndex = 33;
            // 
            // txttxtBarcodeTypePcs
            // 
            this.txttxtBarcodeTypePcs.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txttxtBarcodeTypePcs.Location = new System.Drawing.Point(236, 84);
            this.txttxtBarcodeTypePcs.Name = "txttxtBarcodeTypePcs";
            this.txttxtBarcodeTypePcs.Size = new System.Drawing.Size(427, 43);
            this.txttxtBarcodeTypePcs.TabIndex = 35;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(70, 92);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(130, 35);
            this.label1.TabIndex = 34;
            this.label1.Text = "條碼類型:";
            // 
            // ToolLabel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(707, 622);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Name = "ToolLabel";
            this.Text = "ToolLabel";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtNotPrintLabelNum;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPartNotPrintLabelNum;
        private System.Windows.Forms.Label labInputNun;
        private System.Windows.Forms.TextBox txtKnifeModel_ID;
        private System.Windows.Forms.Label labKnifeModelID;
        private System.Windows.Forms.Button btnPrintLotLabel;
        private System.Windows.Forms.TextBox txtToolNo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnPrintLabel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtBarcodeType;
        private System.Windows.Forms.TextBox txttxtBarcodeTypePcs;
        private System.Windows.Forms.Label label1;
    }
}