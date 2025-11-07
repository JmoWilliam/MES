namespace PrintLabelSystem
{
    partial class OtherCustomerMoldLabel
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtProcess = new System.Windows.Forms.TextBox();
            this.txtBarcodeNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMoId = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(71, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "出貨製程";
            // 
            // txtProcess
            // 
            this.txtProcess.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtProcess.Location = new System.Drawing.Point(173, 42);
            this.txtProcess.Name = "txtProcess";
            this.txtProcess.Size = new System.Drawing.Size(523, 35);
            this.txtProcess.TabIndex = 1;
            this.txtProcess.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProcess_KeyDown);
            // 
            // txtBarcodeNo
            // 
            this.txtBarcodeNo.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBarcodeNo.Location = new System.Drawing.Point(173, 116);
            this.txtBarcodeNo.Name = "txtBarcodeNo";
            this.txtBarcodeNo.Size = new System.Drawing.Size(523, 35);
            this.txtBarcodeNo.TabIndex = 3;
            this.txtBarcodeNo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBarcodeNo_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(71, 119);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(96, 26);
            this.label2.TabIndex = 2;
            this.label2.Text = "模仁條碼";
            // 
            // txtMoId
            // 
            this.txtMoId.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtMoId.Location = new System.Drawing.Point(173, 186);
            this.txtMoId.Name = "txtMoId";
            this.txtMoId.Size = new System.Drawing.Size(523, 35);
            this.txtMoId.TabIndex = 5;
            this.txtMoId.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMoId_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(71, 195);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(96, 26);
            this.label3.TabIndex = 4;
            this.label3.Text = "製令條碼";
            // 
            // OtherCustomerMoldLabel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 274);
            this.Controls.Add(this.txtMoId);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtBarcodeNo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtProcess);
            this.Controls.Add(this.label1);
            this.Name = "OtherCustomerMoldLabel";
            this.Text = "OtherCustomerMoldLabel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtProcess;
        private System.Windows.Forms.TextBox txtBarcodeNo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMoId;
        private System.Windows.Forms.Label label3;
    }
}