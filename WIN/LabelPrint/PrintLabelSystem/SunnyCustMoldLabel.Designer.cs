namespace PrintLabelSystem
{
    partial class SunnyCustMoldLabel
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
            this.TxBarCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txMFG = new System.Windows.Forms.TextBox();
            this.labMFG = new System.Windows.Forms.Label();
            this.txtCheckNi = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // TxBarCode
            // 
            this.TxBarCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxBarCode.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxBarCode.Location = new System.Drawing.Point(197, 142);
            this.TxBarCode.MaxLength = 100;
            this.TxBarCode.Name = "TxBarCode";
            this.TxBarCode.Size = new System.Drawing.Size(464, 36);
            this.TxBarCode.TabIndex = 17;
            this.TxBarCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxBarCode_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(75, 147);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 31);
            this.label1.TabIndex = 16;
            this.label1.Text = "產品條碼:";
            // 
            // txMFG
            // 
            this.txMFG.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txMFG.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txMFG.Location = new System.Drawing.Point(197, 86);
            this.txMFG.MaxLength = 100;
            this.txMFG.Name = "txMFG";
            this.txMFG.Size = new System.Drawing.Size(464, 36);
            this.txMFG.TabIndex = 15;
            this.txMFG.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txMFG_KeyDown);
            // 
            // labMFG
            // 
            this.labMFG.AutoSize = true;
            this.labMFG.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labMFG.Location = new System.Drawing.Point(75, 91);
            this.labMFG.Name = "labMFG";
            this.labMFG.Size = new System.Drawing.Size(116, 31);
            this.labMFG.TabIndex = 14;
            this.labMFG.Text = "製令條碼:";
            // 
            // txtCheckNi
            // 
            this.txtCheckNi.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCheckNi.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtCheckNi.Location = new System.Drawing.Point(197, 36);
            this.txtCheckNi.MaxLength = 100;
            this.txtCheckNi.Name = "txtCheckNi";
            this.txtCheckNi.Size = new System.Drawing.Size(464, 36);
            this.txtCheckNi.TabIndex = 19;
            this.txtCheckNi.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCheckNi_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(27, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(164, 31);
            this.label2.TabIndex = 18;
            this.label2.Text = "是否退鎳重鍍:";
            // 
            // SunnyCustMoldLabel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(715, 214);
            this.Controls.Add(this.txtCheckNi);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TxBarCode);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txMFG);
            this.Controls.Add(this.labMFG);
            this.Name = "SunnyCustMoldLabel";
            this.Text = "SunnyCustMoldLabel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TxBarCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txMFG;
        private System.Windows.Forms.Label labMFG;
        private System.Windows.Forms.TextBox txtCheckNi;
        private System.Windows.Forms.Label label2;
    }
}