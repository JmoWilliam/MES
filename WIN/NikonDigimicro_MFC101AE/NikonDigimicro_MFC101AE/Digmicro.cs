using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace NikonDigimicro_MFC101AE
{
    public partial class Digmicro : Form
    {
        //建立物件實體
        ClassSerialPort _nikonMFC = new ClassSerialPort();

        public object Messagebox { get; private set; }

        public Digmicro()
        {
            InitializeComponent();
            txtBox_Receive.KeyDown += Form1_KeyDown;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //初始化
            _nikonMFC.initial();
            radioButton2.Checked = true;
        }
        private void Digmicro_FormClosing(object sender, FormClosingEventArgs e)
        {
            _nikonMFC.close();
        }
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            //快捷鍵
            if (e.KeyCode == Keys.Space)
            {
                e.Handled = true;
                btnGetValue.PerformClick();
            }
            if (e.KeyCode == Keys.C && e.Control == true)
            {
                e.Handled = true;
                btnConnect.PerformClick();
            }
            if (e.KeyCode == Keys.E && e.Control == true)
            {
                e.Handled = true;
                btnExit.PerformClick();
            }
        }

        //連線
        private void btnConnect_Click(object sender, EventArgs e)
        {
            //初始化
            _nikonMFC.initial(cbx_Port.Text, Convert.ToInt16(cbx_BaudRate.Text), StopBits.Two, 8);

            if (_nikonMFC.open())
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "connect!!" + "   ---" + DateTime.Now.ToLocalTime();
            else
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "Disconnect!!" + "   ---" + DateTime.Now.ToLocalTime();
        }
        //斷線
        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            if (_nikonMFC.close())
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "Disconnect!!" + "   ---" + DateTime.Now.ToLocalTime();
            else
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "Disconnect Failed!!" + "   ---" + DateTime.Now.ToLocalTime();
        }
        //重置
        private void btnReset_Click(object sender, EventArgs e)
        {
            if (_nikonMFC.serialPort.IsOpen)
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + _nikonMFC.sendData("RX\r\n") + "   ---" + DateTime.Now.ToLocalTime();
            else
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + " Please connect again " + "   ---" + DateTime.Now.ToLocalTime();
        }
        //取得高度值
        private void btnGetValue_Click(object sender, EventArgs e)
        {
            string value = "";
            //var rand = new Random(); double randValue = double.Parse("-" + rand.NextDouble().ToString()) + rand.NextDouble(); GetValue_InTable(randValue.ToString());
            if (_nikonMFC.serialPort.IsOpen)
            {
                value = _nikonMFC.sendData("QX\r\n");
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + value + "   ---" + DateTime.Now.ToLocalTime();
                GetValue_InTable(value);
            }
            else
                txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "Please connect again " + "   ---" + DateTime.Now.ToLocalTime();

            txtBox_Receive.Select(txtBox_Receive.Text.Length, 0);
            txtBox_Receive.ScrollToCaret();
            txtBox_Receive.Focus();
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            _nikonMFC.close();
            Close();
        }
        private void btnClear_Click(object sender, EventArgs e)
        {
            txtBox_Receive.Text = string.Empty;
            cbx_Arrange_Type.SelectedIndex = -1;
            cbx_Type.SelectedIndex = -1;
            cbx_Site_Choose.SelectedIndex = -1;
            txtBox_Cave.Text = string.Empty;
            txtBox_Cave_Pcs.Text = string.Empty;
            txtBox_Site.Text = string.Empty;
            txtBox_Site_Repeat.Text = string.Empty;
            txtBox_Order.Text = string.Empty;
            txtBox_ItemName.Text = string.Empty;
            txtBox_Surveyor.Text = string.Empty;
            dataGrid_Arrange.Rows.Clear();
            dataGrid_Arrange.Columns.Clear();
        }

        private void btnSaveTxt_Click(object sender, EventArgs e)
        {
            DateTime dateTime = DateTime.Now.ToLocalTime();

            string myPath = txtBox_Path.Text + "\\" + dateTime.ToString("yyMMdd") + "\\";
            if (!Directory.Exists(myPath))
            {
                Directory.CreateDirectory(myPath);
            }

            if (radioButton1.Checked == true)
            {
                if (txtBox_Receive.Text != String.Empty)
                {
                    if (txtBox_Surveyor.Text != String.Empty)
                    {
                        System.String[] ItemValue = new System.String[5];
                        string fileName = "高度計";
                        ItemValue[0] = cbx_Type.Text;
                        ItemValue[1] = txtBox_ItemName.Text;
                        ItemValue[2] = dateTime.ToString("yyMMdd_HHmm");
                        ItemValue[3] = txtBox_Surveyor.Text;

                        for (int i = 0; i < ItemValue.Length; i++)
                        {
                            if (ItemValue[i] != "")
                                fileName = fileName + "_" + ItemValue[i];
                        }
                        //Pass the filepath and filename to the StreamWriter Constructor
                        if (!File.Exists(myPath + fileName + ".txt"))
                        {
                            StreamWriter sw = new StreamWriter(myPath + fileName + ".txt", false, Encoding.GetEncoding("UTF-8"));
                            //Write a txtReceive all line of text
                            sw.WriteLine(txtBox_Receive.Text);
                            //Close the file
                            sw.Close();
                            txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "txt Save OK." + "   ---" + DateTime.Now.ToLocalTime();
                        }
                        else
                            MessageBox.Show("路徑錯誤！");
                    }
                    else
                        MessageBox.Show("請輸入名稱代號！");
                }
            }
            else if (radioButton2.Checked == true)
            {
                if (txtBox_Surveyor.Text != string.Empty || txtBox_Order.Text != string.Empty)
                {
                    System.String[] ItemValue = new System.String[5];
                    string fileName = "高度計";
                    ItemValue[0] = cbx_Type.Text;
                    ItemValue[1] = txtBox_Order.Text;
                    ItemValue[2] = txtBox_ItemName.Text;
                    ItemValue[3] = dateTime.ToString("yyMMdd_HHmm");
                    ItemValue[4] = txtBox_Surveyor.Text;

                    for (int i = 0; i < ItemValue.Length; i++)
                    {
                        if (ItemValue[i] != "")
                            fileName = fileName + "_" + ItemValue[i];
                    }
                    //Pass the filepath and filename to the StreamWriter Constructor
                    if (!File.Exists(myPath + fileName + ".txt"))
                    {
                        StreamWriter sw = new StreamWriter(myPath + fileName + ".txt", false, Encoding.GetEncoding("Unicode"));

                        string txtHeader = "";
                        string txtData = "";
                        int row = this.dataGrid_Arrange.RowCount - 1;
                        int column = this.dataGrid_Arrange.ColumnCount;
                        for (int i = 0; i < row; i++)
                        {
                            string txtline = "";
                            txtline = dataGrid_Arrange.Rows[i].HeaderCell.Value + "\t";
                            for (int j = 0; j < column; j++)
                            {
                                if (i == 0)
                                    txtHeader = txtHeader + "\t" + dataGrid_Arrange.Columns[j].HeaderText.Trim();

                                if (dataGrid_Arrange[j, i].Value == null)
                                {
                                    txtline = txtline + "\t";
                                }
                                else
                                {
                                    txtline = txtline + dataGrid_Arrange[j, i].Value.ToString() + "\t";
                                }
                            }
                            txtData = txtData + "\r\n" + txtline;

                        }
                        string txtValue = txtHeader + "\r\n" + txtData;
                        sw.Write(txtValue);
                        //close the file
                        sw.Close();
                        txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "txt Save OK." + "   ---" + DateTime.Now.ToLocalTime();
                    }
                    else
                        MessageBox.Show("路徑錯誤！");
                }
                else
                    MessageBox.Show("請輸入名稱代號！");
            }
            txtBox_Receive.Select(txtBox_Receive.Text.Length, 0);
            txtBox_Receive.ScrollToCaret();
            txtBox_Receive.Focus();
        }

        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog path = new FolderBrowserDialog();
            //path.SelectedPath = "\\\\192.168.20.199\\品保跨部門資料夾\\#1.品管\\17.高度計";
            //path.ShowDialog();
            //string myPath = path.SelectedPath + "\\" + dateTime.ToString("yyMMdd") + "\\";

            DateTime dateTime = DateTime.Now.ToLocalTime();
            string myPath = txtBox_Path.Text + dateTime.ToString("yyMMdd") + "\\";

            System.String[] ItemValue = new System.String[5];
            string fileName = "高度計";
            ItemValue[0] = cbx_Type.Text;
            ItemValue[1] = txtBox_Order.Text;
            ItemValue[2] = txtBox_ItemName.Text;
            ItemValue[3] = dateTime.ToString("yyMMdd_HHmm");
            ItemValue[4] = txtBox_Surveyor.Text;

            for (int i = 0; i < ItemValue.Length; i++)
            {
                if (ItemValue[i] != "")
                    fileName = fileName + "_" + ItemValue[i];
            }

            if (!Directory.Exists(myPath))
            {
                Directory.CreateDirectory(myPath);
            }

            Excel.Application excelApp = new Excel.Application
            {
                Visible = false
            };
            if (excelApp == null)
            {
                MessageBox.Show("未能產生Excel物件，可能您的電腦未安裝Excel!");
            }
            else
            {
                Object missing = System.Reflection.Missing.Value;
                excelApp.ScreenUpdating = false;
                excelApp.AskToUpdateLinks = false;
                excelApp.DisplayAlerts = false;
                excelApp.EnableEvents = false;
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet xlsheet = excelWorkbook.Worksheets as Excel.Worksheet;
                xlsheet = (Excel.Worksheet)excelWorkbook.Worksheets[1];
                xlsheet.Name = "DATA";

                if (radioButton1.Checked == true)
                {
                    String line = txtBox_Receive.Text;
                    string[] aryString = (!String.IsNullOrEmpty(txtBox_Receive.Text.Trim())) ? txtBox_Receive.Lines : null;
                    int count = txtBox_Receive.GetLineFromCharIndex(txtBox_Receive.Text.Length);
                    for (int i = 1; i < count; i++)
                    {
                        xlsheet.Cells[i, 1] = aryString[i];
                    }
                    xlsheet.SaveAs(myPath + fileName, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "Excel Save OK." + "   ---" + DateTime.Now.ToLocalTime();
                }
                else if (radioButton2.Checked == true)
                {
                    if (txtBox_Surveyor.Text != string.Empty || txtBox_Order.Text != string.Empty)
                    {
                        //Pass the filepath and filename to the StreamWriter Constructor
                        if (!File.Exists(myPath + fileName + ".xlsx"))
                        {
                            string value = "";
                            int row = this.dataGrid_Arrange.RowCount - 1;
                            int column = this.dataGrid_Arrange.ColumnCount;
                            for (int i = 0; i < row; i++)
                            {
                                xlsheet.Cells[i + 2, 1].value = dataGrid_Arrange.Rows[i].HeaderCell.Value.ToString();
                                xlsheet.Cells[i + 2, 1].NumberFormat = "@";
                                for (int j = 0; j < column; j++)
                                {
                                    xlsheet.Cells[1, j + 2].value = dataGrid_Arrange.Columns[j].HeaderText.Trim();
                                    xlsheet.Cells[1, j + 2].NumberFormat = "@";
                                    value = dataGrid_Arrange.Rows[i].Cells[j].Value.ToString();
                                    xlsheet.Cells[i + 2, j + 2].value = value;
                                }

                            }
                            xlsheet.SaveAs(myPath + fileName, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                            txtBox_Receive.Text = txtBox_Receive.Text.Trim() + "\r\n" + "Excel Save OK." + "   ---" + DateTime.Now.ToLocalTime();
                        }
                        else
                            MessageBox.Show("路徑錯誤！");
                    }
                    else
                        MessageBox.Show("請輸入名稱代號！");
                }

                #region 釋放對象
                // 釋放Workbook對象
                xlsheet = null;
                excelWorkbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                excelWorkbook = null;
                // 釋放Excel對象
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;
                // 調用垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();
                #endregion
            }
            txtBox_Receive.Select(txtBox_Receive.Text.Length, 0);
            txtBox_Receive.ScrollToCaret();
            txtBox_Receive.Focus();
        }

        private void btnArrange_Click(object sender, EventArgs e)
        {
            dataGrid_Arrange.Visible = false;
            int Cave_Quantity = txtBox_Cave.Text != "" ? int.Parse(txtBox_Cave.Text) : 0;
            int CavePcs_Quantity = txtBox_Cave_Pcs.Text != "" ? int.Parse(txtBox_Cave_Pcs.Text) : 0;
            int Site_Quantity = txtBox_Site.Text != "" ? int.Parse(txtBox_Site.Text) : 0;
            int SiteRepeat_Quantity = txtBox_Site_Repeat.Text != "" ? int.Parse(txtBox_Site_Repeat.Text) : 0;
            int column_count = Cave_Quantity * CavePcs_Quantity;
            int row_count = Site_Quantity * SiteRepeat_Quantity;
            int Combo = cbx_Arrange_Type.SelectedIndex;
            string Site = cbx_Site_Choose.Text;
            if (Site == "")
                Site = "尺吋";

            if (column_count * row_count != 0)
            {
                int CaveCount = 1; int CavePcsCount = 1; float sumValue = 0;
                string rowHeader = null;
                int count = txtBox_Receive.GetLineFromCharIndex(txtBox_Receive.Text.Length);
                string[] aryString = (!String.IsNullOrEmpty(txtBox_Receive.Text.Trim())) ? txtBox_Receive.Lines : null;

                //put value
                if (aryString != null)
                {
                    dataGrid_Arrange.Rows.Clear();
                    dataGrid_Arrange.Columns.Clear();
                    // CREATE COLUMNS
                    for (int i = 0; i < column_count; i++)
                    {
                        if (CavePcsCount > CavePcs_Quantity)
                        {
                            CaveCount++;
                            CavePcsCount = 1;
                        }
                        dataGrid_Arrange.Columns.Add("Column" + i, CaveCount + "-" + CavePcsCount);
                        dataGrid_Arrange.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                        CavePcsCount++;

                    }
                    // CREATE ROWS
                    int rowi = 1; int rowj = 1; int rowk = 0;
                    for (int i = 0; i < row_count; i++)
                    {
                        dataGrid_Arrange.Rows.Add();
                        dataGrid_Arrange.Rows[rowk].HeaderCell.Value = Site + " " + rowi + "." + rowj;
                        if (rowj % SiteRepeat_Quantity == 0)
                        {
                            if (cbx_Arrange_Type.Text != "")
                            {
                                if (Combo == 0)
                                {
                                    rowHeader = "最大值";
                                }
                                else if (Combo == 1)
                                {
                                    rowHeader = "平均值";
                                }
                                else if (Combo == 2)
                                {
                                    rowHeader = "平行度";
                                }
                                if (Combo != 3)
                                {
                                    dataGrid_Arrange.Rows.Add();
                                    rowk++;
                                    dataGrid_Arrange.Rows[rowk].HeaderCell.Value = rowHeader;
                                }
                            }
                            rowi++;
                            rowj = 0;
                        }
                        rowk++;
                        rowj++;
                    }

                    // START TO PUT VALUE
                    int row = 0; int col = 0; rowk = 0; float PutValue0 = 0; int comboN = 1; float abcH = 0;
                    for (int i = 0; i < aryString.Length; i++)
                    {
                        if (aryString[i] != null)
                        {
                            if (Regex.IsMatch(aryString[i], @"^[-+][0-9]{4}.[0-9]{4}"))
                            {
                                float PutValue = float.Parse(aryString[i]);
                                dataGrid_Arrange.Rows[rowk].Cells[col].Value = PutValue;
                                if (cbx_Arrange_Type.Text != "")
                                {
                                    if (Combo == 0)
                                    {
                                        //取最大值
                                        if (PutValue >= sumValue)
                                            sumValue = PutValue;
                                    }
                                    else if (Combo == 1)
                                    {
                                        //取平均值前先加總
                                        sumValue = sumValue + PutValue;
                                    }
                                    else if (Combo == 2)
                                    {
                                        //取平行度
                                        if (PutValue != PutValue0)
                                            sumValue = PutValue0;
                                    }
                                    else if (Combo == 3)
                                    {
                                        //黑物abc高度
                                        if (comboN == 1)
                                        {
                                            abcH = PutValue;
                                        }
                                        else
                                        {
                                            dataGrid_Arrange.Rows[rowk].Cells[col].Value = PutValue - abcH;
                                        }
                                    }
                                    comboN++;
                                }
                                row++;
                                rowk++;
                                if (row % SiteRepeat_Quantity == 0)
                                {
                                    if (cbx_Arrange_Type.Text != "")
                                    {
                                        if (Combo == 1)
                                        {
                                            //取平均值
                                            sumValue = sumValue / SiteRepeat_Quantity;
                                        }
                                        else if (Combo == 2)
                                        {
                                            //取平行度
                                            sumValue = (Math.Abs(sumValue) - Math.Abs(PutValue)) * 1000;
                                        }
                                        if (Combo != 3)
                                        {
                                            dataGrid_Arrange.Rows[rowk].Cells[col].Value = sumValue;
                                            rowk++;
                                        }
                                        sumValue = 0;
                                        PutValue0 = 0;
                                        abcH = 0;
                                        comboN = 1;
                                    }

                                }
                                PutValue0 = PutValue;
                                //行數 mod 吋法數*吋法重複數 = 0
                                if (row % (Site_Quantity * SiteRepeat_Quantity) == 0)
                                {
                                    CavePcsCount++;
                                    row = 0;
                                    rowk = 0;
                                    //列數 mod 每穴pcs = 0
                                    if (col % CavePcs_Quantity == 0)
                                    {
                                        CaveCount++;
                                        CavePcsCount = 1;
                                    }
                                    col++;

                                }
                            }
                        }
                    }
                }
                dataGrid_Arrange.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGrid_Arrange.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            dataGrid_Arrange.Visible = true;
        }
        private void GetValue_InTable(string value)
        {
            dataGrid_Arrange.Visible = false;
            startLabel:
            int Cave_Quantity = txtBox_Cave.Text != "" ? int.Parse(txtBox_Cave.Text) : 0;
            int CavePcs_Quantity = txtBox_Cave_Pcs.Text != "" ? int.Parse(txtBox_Cave_Pcs.Text) : 0;
            int Site_Quantity = txtBox_Site.Text != "" ? int.Parse(txtBox_Site.Text) : 0;
            int SiteRepeat_Quantity = txtBox_Site_Repeat.Text != "" ? int.Parse(txtBox_Site_Repeat.Text) : 0;
            int column_count = Cave_Quantity * CavePcs_Quantity;
            int row_count = Site_Quantity * SiteRepeat_Quantity;
            int Combo = cbx_Arrange_Type.SelectedIndex;
            string Site = cbx_Site_Choose.Text;
            if (Site == "")
                Site = "尺吋";

            if (column_count * row_count != 0)
            {
                string rowHeader = null;
                int CaveCount = 1; int CavePcsCount = 1;
                int NowRow = dataGrid_Arrange.RowCount;
                int NowColumn = dataGrid_Arrange.ColumnCount;


                //CREAT COLUMNS AND ROWS
                if (NowColumn < column_count)
                {
                    // CREATE COLUMNS
                    for (int i = 0; i < column_count; i++)
                    {
                        if (CavePcsCount > CavePcs_Quantity)
                        {
                            CaveCount++;
                            CavePcsCount = 1;
                        }
                        dataGrid_Arrange.Columns.Add("Column" + i, CaveCount + "-" + CavePcsCount);
                        dataGrid_Arrange.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                        CavePcsCount++;

                    }
                    if (NowRow < row_count)
                    {
                        // CREATE ROWS
                        int rowi = 1; int rowj = 1; int rowk = 0;
                        for (int i = 0; i < row_count; i++)
                        {
                            dataGrid_Arrange.Rows.Add();
                            dataGrid_Arrange.Rows[rowk].HeaderCell.Value = Site + " " + rowi + "." + rowj;
                            if (rowj % SiteRepeat_Quantity == 0)
                            {
                                if (cbx_Arrange_Type.Text != "")
                                {
                                    if (Combo == 0)
                                    {
                                        rowHeader = "最大值";
                                    }
                                    else if (Combo == 1)
                                    {
                                        rowHeader = "平均值";
                                    }
                                    else if (Combo == 2)
                                    {
                                        rowHeader = "平均值";
                                    }

                                    if (Combo != 3)
                                    {
                                        dataGrid_Arrange.Rows.Add();
                                        rowk++;
                                        dataGrid_Arrange.Rows[rowk].HeaderCell.Value = rowHeader;
                                    }
                                }
                                rowi++;
                                rowj = 0;
                            }
                            rowj++;
                            rowk++;
                        }
                    }
                }

                //START TO PUT VALUE
                // for row

                //排列方法，每吋法+1
                if (cbx_Arrange_Type.Text != "")
                {
                    if (Combo != 3) //abc不加
                    {
                        row_count = row_count + Site_Quantity;
                    }
                    else
                    {
                        if (txtBox_Site_Repeat.Text != "3")
                        {
                            dataGrid_Arrange.Rows.Clear();
                            dataGrid_Arrange.Columns.Clear();
                            txtBox_Site_Repeat.Text = "3";
                            goto startLabel;
                        }
                    }
                }

                int NewRow = 0; int NewColumn = 0; int NewCount = 0;

                //search empty row and column
                for (int i = 0; i < column_count; i++)
                {
                    if (NewCount == 0)
                    {
                        for (int j = 0; j < row_count; j++)
                        {
                            //string dataGridValue = dataGrid_Arrange.Rows[i].Cells[j].Value.ToString();

                            if (dataGrid_Arrange.Rows[j].Cells[i].Value == null)
                            {
                                NewRow = j;
                                NewColumn = i;
                                NewCount++;
                                //MessageBox.Show(NewRow + " " + NewColumn);
                                break;
                            }
                        }
                        
                        //MessageBox.Show(rowsHeader + " " + NewRow + " " + NewColumn);
                    }
                }
                if (Combo == 3)
                {
                    float abcH = 0;
                    //黑物abc
                    if (NewRow % SiteRepeat_Quantity == 1)
                    {
                        abcH = float.Parse(dataGrid_Arrange.Rows[NewRow - 1].Cells[NewColumn].Value.ToString());
                    }
                    else if(NewRow % SiteRepeat_Quantity == 2)
                    {
                        abcH = float.Parse(dataGrid_Arrange.Rows[NewRow - 2].Cells[NewColumn].Value.ToString());
                    }
                    dataGrid_Arrange.Rows[NewRow].Cells[NewColumn].Value = float.Parse(value) - abcH;
                } else
                {
                    dataGrid_Arrange.Rows[NewRow].Cells[NewColumn].Value = float.Parse(value);
                }

                //put 排列方法
                if (cbx_Arrange_Type.Text != "")
                {
                    if (Combo != 3)
                    {
                        string rowsHeader = dataGrid_Arrange.Rows[NewRow + 1].HeaderCell.Value.ToString();
                        if (rowsHeader == "最大值" || rowsHeader == "平均值" || rowsHeader == "平行度")
                        {
                            if (Combo == 0)
                            {
                                float max = 0; float max0 = 0; float maxValue = 0;
                                //最大值
                                for (int k = 0; k < SiteRepeat_Quantity; k++)
                                {
                                    max = float.Parse(dataGrid_Arrange.Rows[NewRow - k].Cells[NewColumn].Value.ToString());

                                    if (max >= max0)
                                        maxValue = max;
                                    else
                                        maxValue = max0;

                                    max0 = max;
                                }
                                dataGrid_Arrange.Rows[NewRow + 1].Cells[NewColumn].Value = maxValue;
                            }
                            else if (Combo == 1)
                            {
                                float sum = 0;
                                //平均值
                                for (int k = 0; k < SiteRepeat_Quantity; k++)
                                {
                                    sum = sum + float.Parse(dataGrid_Arrange.Rows[NewRow - k].Cells[NewColumn].Value.ToString());
                                }
                                dataGrid_Arrange.Rows[NewRow + 1].Cells[NewColumn].Value = sum / SiteRepeat_Quantity;
                            }
                            else if (Combo == 2)
                            {
                                float ValueA = 0; float ValueB = 0; float PutValue = 0;
                                //平行度
                                for (int k = 0; k < SiteRepeat_Quantity; k++)
                                {
                                    ValueA = float.Parse(dataGrid_Arrange.Rows[NewRow - k].Cells[NewColumn].Value.ToString());
                                    if (ValueA != ValueB)
                                    {
                                        PutValue = ValueB;
                                    }
                                    ValueB = ValueA;
                                }
                                PutValue = (Math.Abs(PutValue) - Math.Abs(ValueA)) * 1000;
                                dataGrid_Arrange.Rows[NewRow + 1].Cells[NewColumn].Value = Math.Abs(PutValue);
                            }
                            NewRow++;
                        }
                    }
                }

                dataGrid_Arrange.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGrid_Arrange.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            //endLabel:
            dataGrid_Arrange.Visible = true;
        }
        private void btn_Path_Click(object sender, EventArgs e)
        {
            //選擇資料夾  
            string DefultPath;
            if (txtBox_Path.Text == "")
            {
                DefultPath = "\\\\192.168.20.199\\品保跨部門資料夾\\#1.品管\\17.高度計\\";
            }
            else
            {
                DefultPath = txtBox_Path.Text;
            }
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog
            {
                SelectedPath = DefultPath //預設路徑
            };
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string FolderPath = folderBrowserDialog1.SelectedPath;
                txtBox_Path.Text = FolderPath;
            }
        }
    }
}
