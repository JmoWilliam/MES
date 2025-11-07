namespace NikonDigimicro_MFC101AE
{
    partial class Digmicro
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
            this.btnConnect = new System.Windows.Forms.Button();
            this.btnDisconnect = new System.Windows.Forms.Button();
            this.txtBox_Receive = new System.Windows.Forms.TextBox();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnGetValue = new System.Windows.Forms.Button();
            this.cbx_BaudRate = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cbx_Port = new System.Windows.Forms.ComboBox();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.label15 = new System.Windows.Forms.Label();
            this.btn_Path = new System.Windows.Forms.Button();
            this.txtBox_Path = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtBox_Surveyor = new System.Windows.Forms.TextBox();
            this.txtBox_Order = new System.Windows.Forms.TextBox();
            this.txtBox_ItemName = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.cbx_Type = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.btnSaveExcel = new System.Windows.Forms.Button();
            this.dataGrid_Arrange = new System.Windows.Forms.DataGridView();
            this.btnSaveTxt = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label13 = new System.Windows.Forms.Label();
            this.btnArrange = new System.Windows.Forms.Button();
            this.label17 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbx_Site_Choose = new System.Windows.Forms.ComboBox();
            this.cbx_Arrange_Type = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtBox_Cave = new System.Windows.Forms.TextBox();
            this.txtBox_Site = new System.Windows.Forms.TextBox();
            this.txtBox_Cave_Pcs = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtBox_Site_Repeat = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid_Arrange)).BeginInit();
            this.panel3.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(18, 85);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(75, 23);
            this.btnConnect.TabIndex = 1;
            this.btnConnect.Text = "高度計連線";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // btnDisconnect
            // 
            this.btnDisconnect.Location = new System.Drawing.Point(114, 85);
            this.btnDisconnect.Name = "btnDisconnect";
            this.btnDisconnect.Size = new System.Drawing.Size(75, 23);
            this.btnDisconnect.TabIndex = 3;
            this.btnDisconnect.Text = "斷線";
            this.btnDisconnect.UseVisualStyleBackColor = true;
            this.btnDisconnect.Click += new System.EventHandler(this.btnDisconnect_Click);
            // 
            // txtBox_Receive
            // 
            this.txtBox_Receive.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBox_Receive.Location = new System.Drawing.Point(226, 6);
            this.txtBox_Receive.Multiline = true;
            this.txtBox_Receive.Name = "txtBox_Receive";
            this.txtBox_Receive.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtBox_Receive.Size = new System.Drawing.Size(368, 148);
            this.txtBox_Receive.TabIndex = 2;
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(18, 114);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 25);
            this.btnReset.TabIndex = 3;
            this.btnReset.Text = "高度計重置";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // btnGetValue
            // 
            this.btnGetValue.Location = new System.Drawing.Point(517, 401);
            this.btnGetValue.Name = "btnGetValue";
            this.btnGetValue.Size = new System.Drawing.Size(75, 24);
            this.btnGetValue.TabIndex = 5;
            this.btnGetValue.Text = "擷取";
            this.btnGetValue.UseVisualStyleBackColor = true;
            this.btnGetValue.Click += new System.EventHandler(this.btnGetValue_Click);
            // 
            // cbx_BaudRate
            // 
            this.cbx_BaudRate.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbx_BaudRate.FormattingEnabled = true;
            this.cbx_BaudRate.Items.AddRange(new object[] {
            "4800",
            "9600"});
            this.cbx_BaudRate.Location = new System.Drawing.Point(94, 50);
            this.cbx_BaudRate.Name = "cbx_BaudRate";
            this.cbx_BaudRate.Size = new System.Drawing.Size(95, 21);
            this.cbx_BaudRate.TabIndex = 7;
            this.cbx_BaudRate.TabStop = false;
            this.cbx_BaudRate.Text = "4800";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 9;
            this.label1.Text = "COM Port";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 12);
            this.label2.TabIndex = 10;
            this.label2.Text = "Baud Rate";
            // 
            // cbx_Port
            // 
            this.cbx_Port.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbx_Port.FormattingEnabled = true;
            this.cbx_Port.Items.AddRange(new object[] {
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8"});
            this.cbx_Port.Location = new System.Drawing.Point(94, 14);
            this.cbx_Port.Name = "cbx_Port";
            this.cbx_Port.Size = new System.Drawing.Size(95, 21);
            this.cbx_Port.TabIndex = 11;
            this.cbx_Port.TabStop = false;
            this.cbx_Port.Text = "COM5";
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(524, 12);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 10;
            this.btnExit.Text = "關閉程式";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(114, 115);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 4;
            this.btnClear.Text = "清除全部";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(554, 477);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(40, 12);
            this.label5.TabIndex = 19;
            this.label5.Text = "V 2.0.3";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(3, 80);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(608, 526);
            this.tabControl1.TabIndex = 22;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.radioButton2);
            this.tabPage1.Controls.Add(this.radioButton1);
            this.tabPage1.Controls.Add(this.label15);
            this.tabPage1.Controls.Add(this.btn_Path);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.txtBox_Path);
            this.tabPage1.Controls.Add(this.label12);
            this.tabPage1.Controls.Add(this.panel2);
            this.tabPage1.Controls.Add(this.label16);
            this.tabPage1.Controls.Add(this.btnSaveExcel);
            this.tabPage1.Controls.Add(this.dataGrid_Arrange);
            this.tabPage1.Controls.Add(this.btnSaveTxt);
            this.tabPage1.Controls.Add(this.panel3);
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Controls.Add(this.txtBox_Receive);
            this.tabPage1.Controls.Add(this.btnGetValue);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(600, 500);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Measurement";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(266, 398);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(35, 16);
            this.radioButton2.TabIndex = 37;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "下";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(226, 398);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(35, 16);
            this.radioButton1.TabIndex = 37;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "上";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("新細明體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label15.Location = new System.Drawing.Point(226, 446);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(103, 11);
            this.label15.TabIndex = 30;
            this.label15.Text = "P.S. 人員、製令必填";
            // 
            // btn_Path
            // 
            this.btn_Path.Location = new System.Drawing.Point(43, 465);
            this.btn_Path.Name = "btn_Path";
            this.btn_Path.Size = new System.Drawing.Size(23, 23);
            this.btn_Path.TabIndex = 36;
            this.btn_Path.Text = "...";
            this.btn_Path.UseVisualStyleBackColor = true;
            this.btn_Path.Click += new System.EventHandler(this.btn_Path_Click);
            // 
            // txtBox_Path
            // 
            this.txtBox_Path.Location = new System.Drawing.Point(72, 467);
            this.txtBox_Path.Name = "txtBox_Path";
            this.txtBox_Path.Size = new System.Drawing.Size(299, 22);
            this.txtBox_Path.TabIndex = 35;
            this.txtBox_Path.Text = "\\\\192.168.20.199\\品保跨部門資料夾\\#1.品管\\17.高度計\\";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("新細明體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label12.Location = new System.Drawing.Point(383, 433);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(115, 11);
            this.label12.TabIndex = 22;
            this.label12.Text = "P.S. 有安裝excel才可用";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.txtBox_Surveyor);
            this.panel2.Controls.Add(this.txtBox_Order);
            this.panel2.Controls.Add(this.txtBox_ItemName);
            this.panel2.Controls.Add(this.label14);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.cbx_Type);
            this.panel2.Location = new System.Drawing.Point(8, 335);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(212, 126);
            this.panel2.TabIndex = 29;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(26, 11);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 12);
            this.label6.TabIndex = 21;
            this.label6.Text = "類型：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 37);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 18;
            this.label3.Text = "人員：";
            // 
            // txtBox_Surveyor
            // 
            this.txtBox_Surveyor.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBox_Surveyor.Location = new System.Drawing.Point(88, 34);
            this.txtBox_Surveyor.Name = "txtBox_Surveyor";
            this.txtBox_Surveyor.Size = new System.Drawing.Size(95, 23);
            this.txtBox_Surveyor.TabIndex = 6;
            this.txtBox_Surveyor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtBox_Order
            // 
            this.txtBox_Order.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBox_Order.Location = new System.Drawing.Point(88, 63);
            this.txtBox_Order.Name = "txtBox_Order";
            this.txtBox_Order.Size = new System.Drawing.Size(95, 23);
            this.txtBox_Order.TabIndex = 7;
            this.txtBox_Order.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtBox_ItemName
            // 
            this.txtBox_ItemName.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtBox_ItemName.Location = new System.Drawing.Point(88, 92);
            this.txtBox_ItemName.Name = "txtBox_ItemName";
            this.txtBox_ItemName.Size = new System.Drawing.Size(95, 23);
            this.txtBox_ItemName.TabIndex = 7;
            this.txtBox_ItemName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(26, 66);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(41, 12);
            this.label14.TabIndex = 18;
            this.label14.Text = "製令：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(26, 95);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(41, 12);
            this.label7.TabIndex = 18;
            this.label7.Text = "部番：";
            // 
            // cbx_Type
            // 
            this.cbx_Type.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbx_Type.FormattingEnabled = true;
            this.cbx_Type.Items.AddRange(new object[] {
            "白物鏡片",
            "模造鏡片",
            "黑物",
            "模仁",
            "入料檢",
            "模具",
            "其他"});
            this.cbx_Type.Location = new System.Drawing.Point(88, 8);
            this.cbx_Type.Name = "cbx_Type";
            this.cbx_Type.Size = new System.Drawing.Size(95, 21);
            this.cbx_Type.TabIndex = 11;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(8, 470);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(29, 12);
            this.label16.TabIndex = 34;
            this.label16.Text = "路徑";
            // 
            // btnSaveExcel
            // 
            this.btnSaveExcel.Location = new System.Drawing.Point(385, 402);
            this.btnSaveExcel.Name = "btnSaveExcel";
            this.btnSaveExcel.Size = new System.Drawing.Size(75, 23);
            this.btnSaveExcel.TabIndex = 8;
            this.btnSaveExcel.Text = "儲存excel";
            this.btnSaveExcel.UseVisualStyleBackColor = true;
            this.btnSaveExcel.Click += new System.EventHandler(this.btnSaveExcel_Click);
            // 
            // dataGrid_Arrange
            // 
            this.dataGrid_Arrange.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid_Arrange.Location = new System.Drawing.Point(226, 160);
            this.dataGrid_Arrange.Name = "dataGrid_Arrange";
            this.dataGrid_Arrange.RowTemplate.Height = 24;
            this.dataGrid_Arrange.Size = new System.Drawing.Size(368, 233);
            this.dataGrid_Arrange.TabIndex = 27;
            // 
            // btnSaveTxt
            // 
            this.btnSaveTxt.Location = new System.Drawing.Point(226, 420);
            this.btnSaveTxt.Name = "btnSaveTxt";
            this.btnSaveTxt.Size = new System.Drawing.Size(75, 23);
            this.btnSaveTxt.TabIndex = 9;
            this.btnSaveTxt.Text = "儲存txt";
            this.btnSaveTxt.UseVisualStyleBackColor = true;
            this.btnSaveTxt.Click += new System.EventHandler(this.btnSaveTxt_Click);
            // 
            // panel3
            // 
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.label13);
            this.panel3.Controls.Add(this.btnArrange);
            this.panel3.Controls.Add(this.label17);
            this.panel3.Controls.Add(this.label4);
            this.panel3.Controls.Add(this.cbx_Site_Choose);
            this.panel3.Controls.Add(this.cbx_Arrange_Type);
            this.panel3.Controls.Add(this.label8);
            this.panel3.Controls.Add(this.txtBox_Cave);
            this.panel3.Controls.Add(this.txtBox_Site);
            this.panel3.Controls.Add(this.txtBox_Cave_Pcs);
            this.panel3.Controls.Add(this.label11);
            this.panel3.Controls.Add(this.txtBox_Site_Repeat);
            this.panel3.Controls.Add(this.label9);
            this.panel3.Controls.Add(this.label10);
            this.panel3.Location = new System.Drawing.Point(8, 160);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(212, 169);
            this.panel3.TabIndex = 26;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("新細明體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label13.Location = new System.Drawing.Point(112, 120);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(82, 33);
            this.label13.TabIndex = 31;
            this.label13.Text = "上方欄位填滿後\r\n會自動整理表格\r\n←也可自行整理";
            // 
            // btnArrange
            // 
            this.btnArrange.Location = new System.Drawing.Point(18, 130);
            this.btnArrange.Name = "btnArrange";
            this.btnArrange.Size = new System.Drawing.Size(75, 23);
            this.btnArrange.TabIndex = 27;
            this.btnArrange.Text = "整理表格";
            this.btnArrange.UseVisualStyleBackColor = true;
            this.btnArrange.Click += new System.EventHandler(this.btnArrange_Click);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(7, 75);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(65, 12);
            this.label17.TabIndex = 30;
            this.label17.Text = "量測位置：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 100);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 30;
            this.label4.Text = "排列方法：";
            // 
            // cbx_Site_Choose
            // 
            this.cbx_Site_Choose.FormattingEnabled = true;
            this.cbx_Site_Choose.Items.AddRange(new object[] {
            "高度",
            "肩厚",
            "芯厚"});
            this.cbx_Site_Choose.Location = new System.Drawing.Point(78, 72);
            this.cbx_Site_Choose.Name = "cbx_Site_Choose";
            this.cbx_Site_Choose.Size = new System.Drawing.Size(111, 20);
            this.cbx_Site_Choose.TabIndex = 29;
            // 
            // cbx_Arrange_Type
            // 
            this.cbx_Arrange_Type.FormattingEnabled = true;
            this.cbx_Arrange_Type.Items.AddRange(new object[] {
            "吋法取最大值",
            "吋法取平均值",
            "吋法取平行度",
            "黑物875-ABC高度"});
            this.cbx_Arrange_Type.Location = new System.Drawing.Point(78, 97);
            this.cbx_Arrange_Type.Name = "cbx_Arrange_Type";
            this.cbx_Arrange_Type.Size = new System.Drawing.Size(111, 20);
            this.cbx_Arrange_Type.TabIndex = 29;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(7, 19);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(53, 12);
            this.label8.TabIndex = 18;
            this.label8.Text = "總穴數：";
            // 
            // txtBox_Cave
            // 
            this.txtBox_Cave.Location = new System.Drawing.Point(63, 16);
            this.txtBox_Cave.Name = "txtBox_Cave";
            this.txtBox_Cave.Size = new System.Drawing.Size(31, 22);
            this.txtBox_Cave.TabIndex = 6;
            this.txtBox_Cave.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtBox_Site
            // 
            this.txtBox_Site.Location = new System.Drawing.Point(167, 16);
            this.txtBox_Site.Name = "txtBox_Site";
            this.txtBox_Site.Size = new System.Drawing.Size(31, 22);
            this.txtBox_Site.TabIndex = 6;
            this.txtBox_Site.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtBox_Cave_Pcs
            // 
            this.txtBox_Cave_Pcs.Location = new System.Drawing.Point(63, 44);
            this.txtBox_Cave_Pcs.Name = "txtBox_Cave_Pcs";
            this.txtBox_Cave_Pcs.Size = new System.Drawing.Size(31, 22);
            this.txtBox_Cave_Pcs.TabIndex = 6;
            this.txtBox_Cave_Pcs.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(102, 47);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(65, 12);
            this.label11.TabIndex = 18;
            this.label11.Text = "吋法重複：";
            // 
            // txtBox_Site_Repeat
            // 
            this.txtBox_Site_Repeat.Location = new System.Drawing.Point(167, 44);
            this.txtBox_Site_Repeat.Name = "txtBox_Site_Repeat";
            this.txtBox_Site_Repeat.Size = new System.Drawing.Size(31, 22);
            this.txtBox_Site_Repeat.TabIndex = 6;
            this.txtBox_Site_Repeat.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(7, 47);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(56, 12);
            this.label9.TabIndex = 18;
            this.label9.Text = "每穴pcs：";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(105, 19);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 12);
            this.label10.TabIndex = 18;
            this.label10.Text = "吋法數：";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.cbx_BaudRate);
            this.panel1.Controls.Add(this.btnConnect);
            this.panel1.Controls.Add(this.cbx_Port);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btnClear);
            this.panel1.Controls.Add(this.btnReset);
            this.panel1.Controls.Add(this.btnDisconnect);
            this.panel1.Location = new System.Drawing.Point(8, 6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(212, 148);
            this.panel1.TabIndex = 22;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::NikonDigimicro_MFC101AE.Properties.Resources.JMO;
            this.pictureBox1.Location = new System.Drawing.Point(25, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(259, 50);
            this.pictureBox1.TabIndex = 23;
            this.pictureBox1.TabStop = false;
            // 
            // Digmicro
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(611, 608);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnExit);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.Name = "Digmicro";
            this.Text = "高度計連線程式";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Digmicro_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid_Arrange)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Button btnDisconnect;
        private System.Windows.Forms.TextBox txtBox_Receive;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnGetValue;
        private System.Windows.Forms.ComboBox cbx_BaudRate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbx_Port;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btnArrange;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbx_Arrange_Type;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtBox_Cave;
        private System.Windows.Forms.TextBox txtBox_Site;
        private System.Windows.Forms.TextBox txtBox_Cave_Pcs;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtBox_Site_Repeat;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtBox_Surveyor;
        private System.Windows.Forms.Button btnSaveExcel;
        private System.Windows.Forms.TextBox txtBox_ItemName;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnSaveTxt;
        private System.Windows.Forms.ComboBox cbx_Type;
        private System.Windows.Forms.TextBox txtBox_Order;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.DataGridView dataGrid_Arrange;
        private System.Windows.Forms.Button btn_Path;
        private System.Windows.Forms.TextBox txtBox_Path;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.ComboBox cbx_Site_Choose;
    }
}

