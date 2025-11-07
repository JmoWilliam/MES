namespace PrintLabelSystem
{
    partial class InboundLabel
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
            this.textNgReason = new System.Windows.Forms.TextBox();
            this.lab_NCtitle = new System.Windows.Forms.Label();
            this.btnPRINTER = new System.Windows.Forms.Button();
            this.txtMTL_NAME = new System.Windows.Forms.TextBox();
            this.txtMTL_NO = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cboWAREHOUSE = new System.Windows.Forms.ComboBox();
            this.cboSTATION = new System.Windows.Forms.ComboBox();
            this.txtMFG_ID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.TxBarCode = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtProcessName = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textNgReason
            // 
            this.textNgReason.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.textNgReason.Location = new System.Drawing.Point(141, 305);
            this.textNgReason.Name = "textNgReason";
            this.textNgReason.Size = new System.Drawing.Size(692, 36);
            this.textNgReason.TabIndex = 25;
            // 
            // lab_NCtitle
            // 
            this.lab_NCtitle.AutoSize = true;
            this.lab_NCtitle.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_NCtitle.Location = new System.Drawing.Point(19, 310);
            this.lab_NCtitle.Name = "lab_NCtitle";
            this.lab_NCtitle.Size = new System.Drawing.Size(116, 31);
            this.lab_NCtitle.TabIndex = 24;
            this.lab_NCtitle.Text = "報廢原因:";
            // 
            // btnPRINTER
            // 
            this.btnPRINTER.Location = new System.Drawing.Point(414, 370);
            this.btnPRINTER.Margin = new System.Windows.Forms.Padding(2);
            this.btnPRINTER.Name = "btnPRINTER";
            this.btnPRINTER.Size = new System.Drawing.Size(72, 59);
            this.btnPRINTER.TabIndex = 23;
            this.btnPRINTER.Text = "列印";
            this.btnPRINTER.UseVisualStyleBackColor = true;
            this.btnPRINTER.Click += new System.EventHandler(this.btnPRINTER_Click);
            // 
            // txtMTL_NAME
            // 
            this.txtMTL_NAME.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtMTL_NAME.Location = new System.Drawing.Point(534, 152);
            this.txtMTL_NAME.Margin = new System.Windows.Forms.Padding(2);
            this.txtMTL_NAME.Name = "txtMTL_NAME";
            this.txtMTL_NAME.Size = new System.Drawing.Size(299, 39);
            this.txtMTL_NAME.TabIndex = 22;
            // 
            // txtMTL_NO
            // 
            this.txtMTL_NO.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtMTL_NO.Location = new System.Drawing.Point(138, 153);
            this.txtMTL_NO.Margin = new System.Windows.Forms.Padding(2);
            this.txtMTL_NO.Name = "txtMTL_NO";
            this.txtMTL_NO.Size = new System.Drawing.Size(302, 39);
            this.txtMTL_NO.TabIndex = 21;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(457, 156);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(68, 31);
            this.label5.TabIndex = 20;
            this.label5.Text = "品名:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(70, 160);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(68, 31);
            this.label4.TabIndex = 19;
            this.label4.Text = "品號:";
            // 
            // cboWAREHOUSE
            // 
            this.cboWAREHOUSE.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cboWAREHOUSE.FormattingEnabled = true;
            this.cboWAREHOUSE.Location = new System.Drawing.Point(138, 209);
            this.cboWAREHOUSE.Margin = new System.Windows.Forms.Padding(2);
            this.cboWAREHOUSE.Name = "cboWAREHOUSE";
            this.cboWAREHOUSE.Size = new System.Drawing.Size(695, 38);
            this.cboWAREHOUSE.TabIndex = 18;
            // 
            // cboSTATION
            // 
            this.cboSTATION.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cboSTATION.FormattingEnabled = true;
            this.cboSTATION.Location = new System.Drawing.Point(138, 258);
            this.cboSTATION.Margin = new System.Windows.Forms.Padding(2);
            this.cboSTATION.Name = "cboSTATION";
            this.cboSTATION.Size = new System.Drawing.Size(302, 38);
            this.cboSTATION.TabIndex = 17;
            this.cboSTATION.SelectedIndexChanged += new System.EventHandler(this.cboSTATION_SelectedIndexChanged);
            // 
            // txtMFG_ID
            // 
            this.txtMFG_ID.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtMFG_ID.Location = new System.Drawing.Point(135, 22);
            this.txtMFG_ID.Margin = new System.Windows.Forms.Padding(2);
            this.txtMFG_ID.Name = "txtMFG_ID";
            this.txtMFG_ID.Size = new System.Drawing.Size(698, 39);
            this.txtMFG_ID.TabIndex = 16;
            this.txtMFG_ID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMFG_ID_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(19, 216);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(116, 31);
            this.label3.TabIndex = 15;
            this.label3.Text = "入庫庫別:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(11, 261);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 31);
            this.label2.TabIndex = 14;
            this.label2.Text = "加工階段:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(15, 25);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 31);
            this.label1.TabIndex = 13;
            this.label1.Text = "製令條碼:";
            // 
            // TxBarCode
            // 
            this.TxBarCode.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxBarCode.Location = new System.Drawing.Point(135, 87);
            this.TxBarCode.Margin = new System.Windows.Forms.Padding(2);
            this.TxBarCode.Name = "TxBarCode";
            this.TxBarCode.Size = new System.Drawing.Size(698, 39);
            this.TxBarCode.TabIndex = 27;
            this.TxBarCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxBarCode_KeyDown);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(15, 90);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(116, 31);
            this.label6.TabIndex = 26;
            this.label6.Text = "模仁條碼:";
            // 
            // txtProcessName
            // 
            this.txtProcessName.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtProcessName.Location = new System.Drawing.Point(529, 258);
            this.txtProcessName.Name = "txtProcessName";
            this.txtProcessName.Size = new System.Drawing.Size(304, 36);
            this.txtProcessName.TabIndex = 28;
            // 
            // InboundLabel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(896, 449);
            this.Controls.Add(this.txtProcessName);
            this.Controls.Add(this.TxBarCode);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textNgReason);
            this.Controls.Add(this.lab_NCtitle);
            this.Controls.Add(this.btnPRINTER);
            this.Controls.Add(this.txtMTL_NAME);
            this.Controls.Add(this.txtMTL_NO);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cboWAREHOUSE);
            this.Controls.Add(this.cboSTATION);
            this.Controls.Add(this.txtMFG_ID);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "InboundLabel";
            this.Text = "InboundLabel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textNgReason;
        private System.Windows.Forms.Label lab_NCtitle;
        private System.Windows.Forms.Button btnPRINTER;
        private System.Windows.Forms.TextBox txtMTL_NAME;
        private System.Windows.Forms.TextBox txtMTL_NO;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboWAREHOUSE;
        private System.Windows.Forms.ComboBox cboSTATION;
        private System.Windows.Forms.TextBox txtMFG_ID;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TxBarCode;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtProcessName;
    }
}