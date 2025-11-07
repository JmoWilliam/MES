using MESIII;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NiceLabel.SDK;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintLabelSystem
{
    public partial class SunnyShipLabel : Form
    {
        public static ISDClient clientMes2 = new ISDClient(0);
        private ILabel label;
        public SunnyShipLabel()
        {
            InitializeComponent();
            txtSupplierNo.Text = "A10"; //供應商 A10

            #region//標籤機選擇
            JObject jObj = JObject.FromObject(new
            {
                PARAMETER = JObject.FromObject(new
                {
                    CompanyId = 4
                })
            });
            string jStr = JsonConvert.SerializeObject(jObj);
            DataSet oDS = clientMes2.ctEnumerateData("APISO.QryLabelPrintMachine", jStr);
            string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
            int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
            DataTable dt = oDS.Tables[0].Copy();
            DataRow newRow = dt.NewRow();
            newRow["LabelPrintName"] = "请选择";
            newRow["LabelPrintNo"] = -1; // 假设 -1 代表无效值或默认值
            dt.Rows.InsertAt(newRow, 0);
            cboLabelMachine.DisplayMember = "LabelPrintName";
            cboLabelMachine.ValueMember = "LabelPrintNo";
            cboLabelMachine.DataSource = dt;
            cboLabelMachine.SelectedIndex = 0;
            //https://chatgpt.com/c/f945f423-b5bc-4c2b-8518-dd2c516ea083
            #endregion

            // 設置 TextBox 的 KeyPress 事件處理程序
            txtModelNo.KeyPress += txtModelNo_Press;


            txtProductType.SelectedIndex = 0;
            groupBoxMo.Visible = false;
        }

        private void txtMoId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {                    
                    string MoidText = txtMoId.Text.ToString();
                    int Moid = -1;
                    if (MoidText == "")
                    {
                        throw new Exception("製令不可空值");
                    }
                    else
                    {
                        Moid = int.Parse(MoidText);
                    }
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            Moid
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = clientMes2.ctEnumerateData("APISO.QryManufactureOrder", jStr);
                    string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int nTotalMfgCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                    JObject jResMfg = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = nTotalMfgCount,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    string ErpNo = jResMfg["qrydata"][0]["ErpNo"].ToString();
                    labErpWoNo.Text = ErpNo;
                    string MtlItemNo = jResMfg["qrydata"][0]["MtlItemNo"].ToString();
                    labMtlItemNo.Text = MtlItemNo;
                    string MtlItemName = jResMfg["qrydata"][0]["MtlItemName"].ToString();
                    labMtlItemName.Text = MtlItemName;
                    string MtlItemSpec = jResMfg["qrydata"][0]["MtlItemSpec"].ToString();
                    labMtlItemSpec.Text = MtlItemSpec;
                    string PlanQty = jResMfg["qrydata"][0]["PlanQty"].ToString();
                    labPlanQty.Text = PlanQty;
                    groupBoxMo.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("製令異常:" + ex.ToString());
                }
            }
        }

        private void PrintButton_Click(object sender, EventArgs e)
        {
            try
            {
                string ModelNo="", SunnyName = "", Project = "", Num = "", ProductType = "", Supplier = "", DrawingsNo = "", barcode = "";
                #region //一維條碼 資訊串接
                //模號+部番+方案+數量+類別+供應商+圖號
                if (txtModelNo.Text.ToString() != "")
                {
                    ModelNo = txtModelNo.Text.ToString();
                }
                else { ModelNo = "/"; }

                if (txtSunnyName.Text.ToString() != "")
                {
                    if (txtModelNo.Text.ToString() != "")
                    {
                        SunnyName = "/" + txtSunnyName.Text.ToString();
                    }else
                    {
                        SunnyName = txtSunnyName.Text.ToString();
                    }
                }else { MessageBox.Show("列印異常:【件數】不可空"); }

                if (txtProject.Text.ToString() != "")
                {
                    Project = "/" + txtProject.Text.ToString();
                }else { Project = "/"; }

                if (txtNum.Text.ToString() != "")
                {
                    Num = "/" + txtNum.Text.ToString();
                }else { MessageBox.Show("列印異常:【數量】不可空"); }

                if (txtProductType.Text.ToString() != "")
                {
                    if (txtProductType.Text.ToString() == "A.可动侧模芯" || txtProductType.Text.ToString() == "A")
                    {
                        ProductType = "/A";
                    }else if (txtProductType.Text.ToString() == "B.固定侧模芯" || txtProductType.Text.ToString() == "B")
                    {
                        ProductType = "/B";
                    }else if (txtProductType.Text.ToString() == "C.固定侧衬套" || txtProductType.Text.ToString() == "C")
                    {
                        ProductType = "/C";
                    }else if (txtProductType.Text.ToString() == "D.可动侧衬套" || txtProductType.Text.ToString() == "D")
                    {
                        ProductType = "/D";
                    }else if (txtProductType.Text.ToString() == "E.精定位" || txtProductType.Text.ToString() == "E")
                    {
                        ProductType = "/E";
                    }                    
                }else { ProductType = "/"; }

                if (txtSupplierNo.Text.ToString() != "")
                {
                    Supplier = "/" + txtSupplierNo.Text.ToString();
                }else { Supplier = "/"; }

                if (txtDrawingsNo.Text.ToString() != "")
                {
                    DrawingsNo = "/" + txtDrawingsNo.Text.ToString();
                }else { MessageBox.Show("列印異常:【圖面編號】不可空"); }

                barcode = ModelNo + SunnyName + Project + Num + ProductType + Supplier + DrawingsNo;
                #endregion

                #region //標籤樣式設定+印表機選擇                            
                string LabelPath = "", LabelPathLine;
                string PrintMachine = "";

                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyShipLabel\\LabelName.txt", Encoding.Default);
                while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathLine.ToString();
                }

                PrintMachine = cboLabelMachine.SelectedValue.ToString();
                if (PrintMachine != "-1")
                {
                    PrintMachine = cboLabelMachine.SelectedValue.ToString();
                }
                else
                {
                    StreamReader sr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyShipLabel\\PrinterName.txt", Encoding.Default);
                    String line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        PrintMachine = line.ToString();
                    }
                }

                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;
                #endregion

                //訂單的排定交貨日
                DateTime selectedStartDate = mCalendarPackingDate.SelectionStart;
                string PcPromiseDate = selectedStartDate.ToString("yyyy-MM-dd");

                #region //標籤列印
                label.Variables["BarCode"].SetValue(barcode);
                label.Variables["date"].SetValue(PcPromiseDate);
                string Remarks = txtRemarks.Text.ToString();
                if (Remarks == "")
                {
                    label.Variables["note"].SetValue("晶彩");
                }
                else {
                    label.Variables["note"].SetValue(Remarks);
                }
                
                label.Print(1);
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show("舜宇外箱標籤列印異常:" + ex.ToString());
            }
        }

        private void txtModelNo_Press(object sender, KeyPressEventArgs e)
        {
            // 允許空值（backspace）和特定符號
           if (char.IsControl(e.KeyChar) || e.KeyChar == '-')
            {
                return;
            }

            // 允許大寫字母、數字和特定符號
            if (char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '-')
            {
                // 將小寫字母轉換為大寫
                e.KeyChar = char.ToUpper(e.KeyChar);
                return;
            }

            // 阻止其他字符輸入
            e.Handled = true;
        }

        private void txtSunnyName_Press(object sender, KeyPressEventArgs e)
        {
            // 允許空值（backspace）和特定符號
            if (char.IsControl(e.KeyChar) || e.KeyChar == '-')
            {
                return;
            }

            // 允許大寫字母、數字和特定符號
            if (char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '-')
            {
                // 將小寫字母轉換為大寫
                e.KeyChar = char.ToUpper(e.KeyChar);
                return;
            }

            // 阻止其他字符輸入
            e.Handled = true;
        }

        private void txtProject_Press(object sender, KeyPressEventArgs e)
        {
            // 允許空值（backspace）和特定符號
            if (char.IsControl(e.KeyChar) || e.KeyChar == '-')
            {
                return;
            }

            // 允許大寫字母、數字和特定符號
            if (char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '-')
            {
                // 將小寫字母轉換為大寫
                e.KeyChar = char.ToUpper(e.KeyChar);
                return;
            }

            // 阻止其他字符輸入
            e.Handled = true;
        }

        private void txtNum_Press(object sender, KeyPressEventArgs e)
        {
            // 允許控制字符（如 backspace）
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // 允許數字輸入
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // 阻止其他字符輸入
            e.Handled = true;
        }

        private void txtDrawingsNo_Press(object sender, KeyPressEventArgs e)
        {
            // 允許空值（backspace）和特定符號
            if (char.IsControl(e.KeyChar) || e.KeyChar == '-')
            {
                return;
            }

            // 允許大寫字母、數字和特定符號
            if (char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '-')
            {
                // 將小寫字母轉換為大寫
                e.KeyChar = char.ToUpper(e.KeyChar);
                return;
            }

            // 阻止其他字符輸入
            e.Handled = true;
        }
    }
}
