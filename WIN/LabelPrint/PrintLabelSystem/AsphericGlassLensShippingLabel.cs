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
    public partial class AsphericGlassLensShippingLabel : Form
    {
        private ILabel label;
        public AsphericGlassLensShippingLabel()
        {
            InitializeComponent();

            gbCustInfo.Visible = false;
            gbCoatingInfo.Visible = false;
            gbLabelInfo.Visible = false;

            btnCustInfoFinish.Visible = false;
            btnLabel.Visible = false;
        }

        private void btnBarcode_Click(object sender, EventArgs e)
        {
            try
            {
                textPrintNum.Text = "1";

                gbCustInfo.Visible = true;
                btnCustInfoFinish.Visible = true;


                string BARCODE_NO = txtBarcodeNo.Text.ToString().ToUpper();
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        BARCODE_NO
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryBarcodeAttribute", jStr);
                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = TOTAL_COUNT,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });
                string MTL_NO = jRes["qrydata"][0]["MtlItemNo"].ToString();
                string CustomerMtlNo="", CustomerMtlName = "", CustomerNo = "", CustomerPartNo = "";
                if (MTL_NO == "3OLD23XXXXX-C3-61")
                {
                    CustomerMtlNo = "M022C02009011"; //客戶品號
                    CustomerMtlName = "D23-C3"; //客戶品名
                    CustomerNo = "E005"; //供應商代碼
                    CustomerPartNo = "23.99J16G003 B"; //客戶料號
                }
                else {
                    #region //查詢ERP 品號jRes，品名，客戶品號，客戶品名                
                    jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            MTL_NO
                        })
                    });
                    jStr = JsonConvert.SerializeObject(jObj);
                    oDS = MESConfig.clientErp.ctEnumerateData("APISO.QryCustomerItem01", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    JObject jResErp = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = TOTAL_COUNT,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    CustomerMtlNo = jResErp["qrydata"][0]["MG003"].ToString(); //客戶品號
                    CustomerMtlName = jResErp["qrydata"][0]["MG005"].ToString(); //客戶品名
                    CustomerNo = jResErp["qrydata"][0]["MG001"].ToString(); //供應商代碼
                    CustomerPartNo = jResErp["qrydata"][0]["MG006"].ToString(); //客戶料號
                }

                txtCustomerMtlNo.Text = CustomerMtlNo; //客戶品號
                txtCustomerMtlName.Text = CustomerMtlName; //客戶品名
                txtCustomerNo.Text = ""; //供應商代碼 -- ERP有數值也不帶值
                txtCustomerPartNo.Text = CustomerPartNo;//客戶料號
                #endregion
            }
            catch (Exception ex)
            {
                MESConfig.LogData("鍍膜塑膠鏡片標籤:" + ex.ToString());
            }
        }

        private void btnCustInfoFinish_Click(object sender, EventArgs e)
        {
            try
            {
                gbCoatingInfo.Visible = true;
                btnLabel.Visible = true;
                string BARCODE_NO = txtBarcodeNo.Text.ToString().ToUpper();
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        BARCODE_NO
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryBarcodeAttribute", jStr);
                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = TOTAL_COUNT,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });

                txPot.Text = jRes["qrydata"][0]["PotNumberItemValue"].ToString();
                txCavityNo.Text = jRes["qrydata"][0]["CavityItemValue"].ToString();
                dtCoatingDate.Value = DateTime.Parse(jRes["qrydata"][0]["PotNumberFinishDate"].ToString());
            }
            catch (Exception ex)
            {
                MESConfig.LogData("鍍膜塑膠鏡片標籤:" + ex.ToString());
            }
        }

        private void btnLabel_Click(object sender, EventArgs e)
        {
            try
            {
                gbLabelInfo.Visible = true;
               

                #region//帶出 標籤內容
                string year = YaerMapNo(dtCoatingDate.Value.Year.ToString());
                string month = MpnthMapNo(dtCoatingDate.Value.Month.ToString());
                string day = dtCoatingDate.Value.Day.ToString().PadLeft(2, '0');
                string potno = PotMapNo(txPot.Text.Split('-')[txPot.Text.Split('-').Length - 1].ToString());
                int cavityNo = int.Parse(txCavityNo.Text.ToString());
                string UmbrellaStand = txtUmbrellaStand.Text.ToString().Replace(" ", "");
                string RoundCount = txtRoundCount.Text.ToString().Replace(" ", "");

                if (UmbrellaStand.Length == 0) {
                    UmbrellaStand = "";
                }
                else {
                    UmbrellaStand = "_"+ UmbrellaStand;
                }

                if (RoundCount.Length == 0) {
                    RoundCount = "";
                }

                string LabelContext = txtCustomerPartNo.Text + potno.Replace(" ", "") + year.Replace(" ", "") + month.Replace(" ", "") + day.Replace(" ", "")+ UmbrellaStand.Replace("_", "")+ RoundCount;



                string insideLabelContext = txtCustomerMtlName.Text + "_" + dtCoatingDate.Value.Year.ToString() + "_" + dtCoatingDate.Value.Month.ToString().PadLeft(2, '0') + dtCoatingDate.Value.Day.ToString().PadLeft(2, '0') + "_" + txPot.Text.Split('-')[txPot.Text.Split('-').Length - 1].ToString() + UmbrellaStand+ RoundCount;

                txLabelContext.Text = LabelContext.Replace(" ", "");
                txtinsideLabelContext.Text= insideLabelContext.Replace(" ", "");
                #endregion


            }
            catch (Exception ex)
            {
                MESConfig.LogData("鍍膜塑膠鏡片標籤:" + ex.ToString());
            }
        }
        public string YaerMapNo(string year)
        {
            string[] List = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            int i;
            if (int.Parse(year) < 2020)
            {
                return "Error";
            }
            else
            {
                i = int.Parse(year) - 2020;
                return List[i].ToString();
            }
        }
        public string MpnthMapNo(string month)
        {
            string[] List = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C" };
            return List[int.Parse(month) - 1].ToString();
        }
        public string PotMapNo(string potno)
        {
            string[] List = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            int i;
            if ((int.Parse(potno) - 1) < 10)
            {
                return List[int.Parse(potno) - 1].ToString();
            }
            else
            {
                i = int.Parse(potno) - 10;
                return List[i].ToString();
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            try
            {
                int PrintNum = int.Parse(textPrintNum.Text.ToString());
                if(PrintNum>0) 
                {
                    #region//列印標籤
                    #region //標籤樣式設定+印表機選擇                            
                    string LabelPath = "", LabelPathLine;
                    string PrintMachine = "", PrintMachineLine;

                    StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\AsphericGlassLensShippingLabel\\LabelName.txt", Encoding.Default);
                    while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
                    {
                        LabelPath = LabelPathLine.ToString();
                    }

                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\AsphericGlassLensShippingLabel\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineLine = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineLine.ToString();
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

                    label.Variables["CONTEXT"].SetValue(txLabelContext.Text.ToString());
                    label.Variables["INSIDE_CONTEXT"].SetValue(txtinsideLabelContext.Text.ToString());
                    if (PrintNum > 1)
                    {
                        label.Print(PrintNum);
                    }
                    else
                    {
                        label.Print(1);
                    }
                    #endregion

                }
                else{
                    MessageBox.Show("列印次數不可小於等於 0 ");
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show("鍍膜塑膠鏡片標籤:" + ex.ToString());
            }
        }

        private void textPrintNumKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    int PrintNum = int.Parse(textPrintNum.Text.ToString());
                    if (PrintNum > 0)
                    {

                        #region//列印標籤
                        #region //標籤樣式設定+印表機選擇                            
                        string LabelPath = "", LabelPathLine;
                        string PrintMachine = "", PrintMachineLine;

                        StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\AsphericGlassLensShippingLabel\\LabelName.txt", Encoding.Default);
                        while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
                        {
                            LabelPath = LabelPathLine.ToString();
                        }

                        StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\AsphericGlassLensShippingLabel\\PrinterName.txt", Encoding.Default);
                        while ((PrintMachineLine = PrintMachineTxt.ReadLine()) != null)
                        {
                            PrintMachine = PrintMachineLine.ToString();
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

                        label.Variables["CONTEXT"].SetValue(txLabelContext.Text.ToString());
                        label.Variables["INSIDE_CONTEXT"].SetValue(txtinsideLabelContext.Text.ToString());
                        if (PrintNum>1) {
                            label.Print(PrintNum);
                        }
                        else {
                            label.Print(1);
                        }
                        
                        #endregion

                    }
                    else
                    {
                        MessageBox.Show("列印次數不可小於等於 0 ");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("鍍膜塑膠鏡片標籤:" + ex.ToString());
                }
            }
        }
    }
}