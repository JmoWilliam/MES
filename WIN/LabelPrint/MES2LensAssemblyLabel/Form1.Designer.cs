namespace MES2LensAssemblyLabel
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
            this.cboLabelType = new System.Windows.Forms.ComboBox();
            this.labMFG = new System.Windows.Forms.Label();
            this.txMFG = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txLotBarcode = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(158, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 31);
            this.label2.TabIndex = 13;
            this.label2.Text = "標籤";
            // 
            // cboLabelType
            // 
            this.cboLabelType.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cboLabelType.FormattingEnabled = true;
            this.cboLabelType.Location = new System.Drawing.Point(248, 50);
            this.cboLabelType.Name = "cboLabelType";
            this.cboLabelType.Size = new System.Drawing.Size(660, 39);
            this.cboLabelType.TabIndex = 14;
            // 
            // labMFG
            // 
            this.labMFG.AccessibleRole = System.Windows.Forms.AccessibleRole.PageTabList;
            this.labMFG.AutoSize = true;
            this.labMFG.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labMFG.Location = new System.Drawing.Point(104, 172);
            this.labMFG.Name = "labMFG";
            this.labMFG.Size = new System.Drawing.Size(116, 31);
            this.labMFG.TabIndex = 15;
            this.labMFG.Text = "製令條碼:";
            // 
            // txMFG
            // 
            this.txMFG.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txMFG.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txMFG.Location = new System.Drawing.Point(248, 167);
            this.txMFG.MaxLength = 100;
            this.txMFG.Name = "txMFG";
            this.txMFG.Size = new System.Drawing.Size(660, 36);
            this.txMFG.TabIndex = 16;
            this.txMFG.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txMFG_KeyDown);
            // 
            // label1
            // 
            this.label1.AccessibleRole = System.Windows.Forms.AccessibleRole.PageTabList;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(104, 285);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 31);
            this.label1.TabIndex = 17;
            this.label1.Text = "批號條碼:";
            // 
            // txLotBarcode
            // 
            this.txLotBarcode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txLotBarcode.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txLotBarcode.Location = new System.Drawing.Point(248, 280);
            this.txLotBarcode.MaxLength = 100;
            this.txLotBarcode.Name = "txLotBarcode";
            this.txLotBarcode.Size = new System.Drawing.Size(660, 36);
            this.txLotBarcode.TabIndex = 18;
            this.txLotBarcode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txLotBarcode_KeyDown);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(980, 374);
            this.Controls.Add(this.txLotBarcode);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txMFG);
            this.Controls.Add(this.labMFG);
            this.Controls.Add(this.cboLabelType);
            this.Controls.Add(this.label2);
            this.Name = "Form1";
            this.Text = "【MES 2.0】組立批量標籤";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboLabelType;
        private System.Windows.Forms.Label labMFG;
        private System.Windows.Forms.TextBox txMFG;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txLotBarcode;
    }
}

