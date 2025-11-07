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
    public partial class OtherCustomerMoldLabel : Form
    {
        private ILabel label;
        public OtherCustomerMoldLabel()
        {
            InitializeComponent();
        }

        private void txtProcess_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                switch (txtProcess.Text.ToString()) {
                    case "1001":
                    case "粗銑":
                        txtProcess.Text = "粗銑";
                        break;
                    case "1002":
                    case "鍍鎳":
                        txtProcess.Text = "鍍鎳";
                        break;
                    case "1003":
                    case "開粗外徑":
                        txtProcess.Text = "開粗外徑";
                        break;
                    case "1004":
                    case "咬降":
                        txtProcess.Text = "咬降";
                        break;
                    case "1005":
                    case "精銑":
                        txtProcess.Text = "精銑";
                        break;
                    case "1006":
                    case "精車前一刀":
                        txtProcess.Text = "精車前一刀";
                        break;
                    case "1007":
                    case "精車":
                        txtProcess.Text = "精車";
                        break;
                    case "1008":
                    case "移模備品":
                        txtProcess.Text = "移模備品";
                        break;
                    case "1009":
                    case "研磨出貨":
                        txtProcess.Text = "研磨出貨";
                        break;

                }
                txtBarcodeNo.Focus();
            }
        }

        private void txtBarcodeNo_KeyDown(object sender, KeyEventArgs e)
        {
            try {
                if (e.KeyCode == Keys.Enter)
                {
                    #region //
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            BarcodeNo = txtBarcodeNo.Text.ToString(),
                            MoId = "-1"
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryOtherCustomerMold", jStr);
                    string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int Count = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    JObject jRes = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = Count,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    #endregion
                    for (int i=0;i< Count;i++)
                    {
                        PrintLabel(jRes["qrydata"][i]["MtlItemName"].ToString(), jRes["qrydata"][i]["BarcodeNo"].ToString(), txtProcess.Text.ToString(), jRes["qrydata"][i]["ItemValue"].ToString());
                    }
                    check();
                }
            }
            catch (Exception ex)
            {
                MESConfig.LogData("其他出貨標籤:" + ex.ToString());
                MessageBox.Show("其他出貨標籤:" + ex.ToString());
            }

        }

        private void txtMoId_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    #region //查詢工具型號與明細
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            BarcodeNo = "",
                            MoId = txtMoId.Text.ToString()
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryOtherCustomerMold", jStr);
                    string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int Count = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    JObject jRes = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = Count,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    #endregion
                    for (int i = 0; i < Count; i++)
                    {
                        PrintLabel(jRes["qrydata"][i]["MtlItemName"].ToString(), jRes["qrydata"][i]["BarcodeNo"].ToString(), txtProcess.Text.ToString(), jRes["qrydata"][i]["ItemValue"].ToString());
                    }
                    check();
                }
            }
            catch (Exception ex)
            {
                MESConfig.LogData("其他出貨標籤:" + ex.ToString());
                MessageBox.Show("其他出貨標籤:" + ex.ToString());
            }
        }

        public void PrintLabel(string MtlItemName, string BarcodeNo,string Process,string ItemValue) {
            #region //標籤樣式設定+印表機選擇                            
            string LabelPath = "", LabelPathLine;
            string PrintMachine = "", PrintMachineLine;

            StreamReader LabelPathTxt = new StreamReader(@"C:\WIN\PrintLabelSystem\Template\OtherCustomerMoldLabel\LabelName.txt", Encoding.Default);
            while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
            {
                LabelPath = LabelPathLine.ToString();
            }

            StreamReader PrintMachineTxt = new StreamReader(@"C:\WIN\PrintLabelSystem\Template\OtherCustomerMoldLabel\PrinterName.txt", Encoding.Default);
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

            //取得流水碼
            string label_name = ItemValue.Substring(ItemValue.Length - 2);
            //部番
            string[] WIP_NAME = MtlItemName.Split(' ');
            label.Variables["WIP_NAME"].SetValue(WIP_NAME[1] + '-' + label_name);
            label.Variables["BARCODE_NO"].SetValue(BarcodeNo);
            label.Variables["STATION"].SetValue(Process);
            label.Print(1);            
        }

        private void check()
        {
            txtMoId.Text = "";
            txtBarcodeNo.Text = "";
        }
    }
}
