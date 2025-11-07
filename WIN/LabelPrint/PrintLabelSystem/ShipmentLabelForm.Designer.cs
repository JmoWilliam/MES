namespace PrintLabelSystem
{
    partial class ShipmentLabelForm
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
            this.label2 = new System.Windows.Forms.Label();
            this.txtLabelFlowID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtLabelFormat = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPiece = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cboLabelType = new System.Windows.Forms.ComboBox();
            this.cboVendor = new System.Windows.Forms.ComboBox();
            this.btnPrint = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(180, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 30);
            this.label1.TabIndex = 0;
            this.label1.Text = "標籤類別";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(228, 131);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 30);
            this.label2.TabIndex = 2;
            this.label2.Text = "廠商";
            // 
            // txtLabelFlowID
            // 
            this.txtLabelFlowID.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtLabelFlowID.Location = new System.Drawing.Point(295, 201);
            this.txtLabelFlowID.Name = "txtLabelFlowID";
            this.txtLabelFlowID.Size = new System.Drawing.Size(296, 39);
            this.txtLabelFlowID.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(156, 204);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(133, 30);
            this.label3.TabIndex = 4;
            this.label3.Text = "目前流水號";
            // 
            // txtLabelFormat
            // 
            this.txtLabelFormat.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtLabelFormat.Location = new System.Drawing.Point(295, 278);
            this.txtLabelFormat.Name = "txtLabelFormat";
            this.txtLabelFormat.Size = new System.Drawing.Size(296, 39);
            this.txtLabelFormat.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(228, 281);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(61, 30);
            this.label4.TabIndex = 6;
            this.label4.Text = "格式";
            // 
            // txtPiece
            // 
            this.txtPiece.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtPiece.Location = new System.Drawing.Point(295, 355);
            this.txtPiece.Name = "txtPiece";
            this.txtPiece.Size = new System.Drawing.Size(296, 39);
            this.txtPiece.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(228, 358);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(61, 30);
            this.label5.TabIndex = 8;
            this.label5.Text = "張數";
            // 
            // cboLabelType
            // 
            this.cboLabelType.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cboLabelType.FormattingEnabled = true;
            this.cboLabelType.Location = new System.Drawing.Point(295, 54);
            this.cboLabelType.Name = "cboLabelType";
            this.cboLabelType.Size = new System.Drawing.Size(296, 38);
            this.cboLabelType.TabIndex = 10;
            // 
            // cboVendor
            // 
            this.cboVendor.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cboVendor.FormattingEnabled = true;
            this.cboVendor.Location = new System.Drawing.Point(295, 131);
            this.cboVendor.Name = "cboVendor";
            this.cboVendor.Size = new System.Drawing.Size(296, 38);
            this.cboVendor.TabIndex = 11;
            this.cboVendor.SelectedIndexChanged += new System.EventHandler(this.cboVendor_SelectedIndexChanged);
            // 
            // btnPrint
            // 
            this.btnPrint.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnPrint.Location = new System.Drawing.Point(377, 429);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(81, 37);
            this.btnPrint.TabIndex = 12;
            this.btnPrint.Text = "列印";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // ShipmentLabelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 499);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.cboVendor);
            this.Controls.Add(this.cboLabelType);
            this.Controls.Add(this.txtPiece);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtLabelFormat);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtLabelFlowID);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "ShipmentLabelForm";
            this.Text = "ShipmentLabelForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLabelFlowID;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtLabelFormat;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPiece;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboLabelType;
        private System.Windows.Forms.ComboBox cboVendor;
        private System.Windows.Forms.Button btnPrint;
    }
}