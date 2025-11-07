namespace PrintLabelSystem
{
    partial class NewMaxCustomerMoldLabel
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
            this.txMFG_ID = new System.Windows.Forms.TextBox();
            this.labMFG_ID = new System.Windows.Forms.Label();
            this.text_InsertType = new System.Windows.Forms.TextBox();
            this.label_InsertType = new System.Windows.Forms.Label();
            this.TxTYPE = new System.Windows.Forms.TextBox();
            this.labType = new System.Windows.Forms.Label();
            this.TxBarCode = new System.Windows.Forms.TextBox();
            this.labBarCode = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txMFG_ID
            // 
            this.txMFG_ID.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txMFG_ID.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txMFG_ID.Location = new System.Drawing.Point(220, 156);
            this.txMFG_ID.MaxLength = 100;
            this.txMFG_ID.Name = "txMFG_ID";
            this.txMFG_ID.Size = new System.Drawing.Size(488, 36);
            this.txMFG_ID.TabIndex = 32;
            this.txMFG_ID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txMFG_ID_KeyDown);
            // 
            // labMFG_ID
            // 
            this.labMFG_ID.AutoSize = true;
            this.labMFG_ID.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labMFG_ID.Location = new System.Drawing.Point(98, 161);
            this.labMFG_ID.Name = "labMFG_ID";
            this.labMFG_ID.Size = new System.Drawing.Size(116, 31);
            this.labMFG_ID.TabIndex = 31;
            this.labMFG_ID.Text = "製令條碼:";
            // 
            // text_InsertType
            // 
            this.text_InsertType.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.text_InsertType.Location = new System.Drawing.Point(220, 307);
            this.text_InsertType.Name = "text_InsertType";
            this.text_InsertType.Size = new System.Drawing.Size(488, 39);
            this.text_InsertType.TabIndex = 30;
            this.text_InsertType.KeyDown += new System.Windows.Forms.KeyEventHandler(this.text_InsertType_KeyDown);
            // 
            // label_InsertType
            // 
            this.label_InsertType.AutoSize = true;
            this.label_InsertType.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_InsertType.Location = new System.Drawing.Point(98, 315);
            this.label_InsertType.Name = "label_InsertType";
            this.label_InsertType.Size = new System.Drawing.Size(116, 31);
            this.label_InsertType.TabIndex = 29;
            this.label_InsertType.Text = "入子類型:";
            // 
            // TxTYPE
            // 
            this.TxTYPE.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxTYPE.Location = new System.Drawing.Point(220, 225);
            this.TxTYPE.Margin = new System.Windows.Forms.Padding(2);
            this.TxTYPE.Name = "TxTYPE";
            this.TxTYPE.ReadOnly = true;
            this.TxTYPE.Size = new System.Drawing.Size(488, 36);
            this.TxTYPE.TabIndex = 28;
            // 
            // labType
            // 
            this.labType.AutoSize = true;
            this.labType.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labType.Location = new System.Drawing.Point(146, 230);
            this.labType.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labType.Name = "labType";
            this.labType.Size = new System.Drawing.Size(68, 31);
            this.labType.TabIndex = 27;
            this.labType.Text = "類別:";
            // 
            // TxBarCode
            // 
            this.TxBarCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxBarCode.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TxBarCode.Location = new System.Drawing.Point(220, 78);
            this.TxBarCode.MaxLength = 100;
            this.TxBarCode.Name = "TxBarCode";
            this.TxBarCode.Size = new System.Drawing.Size(488, 36);
            this.TxBarCode.TabIndex = 26;
            this.TxBarCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxBarCode_KeyDown);
            // 
            // labBarCode
            // 
            this.labBarCode.AccessibleRole = System.Windows.Forms.AccessibleRole.PageTabList;
            this.labBarCode.AutoSize = true;
            this.labBarCode.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labBarCode.Location = new System.Drawing.Point(98, 83);
            this.labBarCode.Name = "labBarCode";
            this.labBarCode.Size = new System.Drawing.Size(116, 31);
            this.labBarCode.TabIndex = 25;
            this.labBarCode.Text = "模仁條碼:";
            // 
            // NewMaxCustomerMoldLabel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.txMFG_ID);
            this.Controls.Add(this.labMFG_ID);
            this.Controls.Add(this.text_InsertType);
            this.Controls.Add(this.label_InsertType);
            this.Controls.Add(this.TxTYPE);
            this.Controls.Add(this.labType);
            this.Controls.Add(this.TxBarCode);
            this.Controls.Add(this.labBarCode);
            this.Name = "NewMaxCustomerMoldLabel";
            this.Text = "NewMaxCustomerMoldLabel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txMFG_ID;
        private System.Windows.Forms.Label labMFG_ID;
        private System.Windows.Forms.TextBox text_InsertType;
        private System.Windows.Forms.Label label_InsertType;
        private System.Windows.Forms.TextBox TxTYPE;
        private System.Windows.Forms.Label labType;
        private System.Windows.Forms.TextBox TxBarCode;
        private System.Windows.Forms.Label labBarCode;
    }
}