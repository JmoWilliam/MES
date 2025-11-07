namespace WhiteLabelPrint
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label2 = new System.Windows.Forms.Label();
            this.TxMFG_ID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TxBarcodeNo = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(11, 67);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(130, 35);
            this.label2.TabIndex = 2;
            this.label2.Text = "製令條碼:";
            // 
            // TxMFG_ID
            // 
            this.TxMFG_ID.Font = new System.Drawing.Font("微軟正黑體", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxMFG_ID.Location = new System.Drawing.Point(145, 59);
            this.TxMFG_ID.Margin = new System.Windows.Forms.Padding(2);
            this.TxMFG_ID.Name = "TxMFG_ID";
            this.TxMFG_ID.Size = new System.Drawing.Size(269, 43);
            this.TxMFG_ID.TabIndex = 4;
            this.TxMFG_ID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxMFG_ID_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(11, 174);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(130, 35);
            this.label3.TabIndex = 5;
            this.label3.Text = "批號條碼:";
            // 
            // TxBarcodeNo
            // 
            this.TxBarcodeNo.Font = new System.Drawing.Font("微軟正黑體", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxBarcodeNo.Location = new System.Drawing.Point(145, 166);
            this.TxBarcodeNo.Margin = new System.Windows.Forms.Padding(2);
            this.TxBarcodeNo.Name = "TxBarcodeNo";
            this.TxBarcodeNo.Size = new System.Drawing.Size(269, 43);
            this.TxBarcodeNo.TabIndex = 6;
            this.TxBarcodeNo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxBarcodeNo_KeyDown);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(513, 283);
            this.Controls.Add(this.TxBarcodeNo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.TxMFG_ID);
            this.Controls.Add(this.label2);
            this.Name = "Form1";
            this.Text = "【MES 2.0】白物標籤列印";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TxMFG_ID;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TxBarcodeNo;
    }
}

