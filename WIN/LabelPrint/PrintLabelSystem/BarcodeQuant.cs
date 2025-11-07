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
    public partial class BarcodeQuant : Form
    {
        public string Data { get; set; }
        private ILabel label;
        public static ISDClient clientMes2 = new ISDClient(0);
        string PrinterName = "", LabelName = "";
        DataSet oDS; string sData = "";
        public BarcodeQuant()
        {
            InitializeComponent();

            StreamReader sr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\BarcodeQuant\\PrinterName.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                PrinterName = line.ToString();
            }
            StreamReader txLabelNamesr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\BarcodeQuant\\LabelName.txt", Encoding.Default);
            String lineLabelName;
            while ((lineLabelName = txLabelNamesr.ReadLine()) != null)
            {
                LabelName = lineLabelName.ToString();
            }
        }

        protected override void OnCreateControl()
        {
            this.InitializePrintEngine();
            base.OnCreateControl();
        }
        protected override void OnClosed(EventArgs e)
        {
            PrintEngineFactory.PrintEngine.Shutdown();
            base.OnClosed(e);
        }
        private void InitializePrintEngine()
        {
            try
            {
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
            }
            catch (SDKException exception)
            {
                MessageBox.Show("Initialization of the SDK failed." + Environment.NewLine + Environment.NewLine + exception.ToString());
                Application.Exit();
            }
        }

        private void txtBarcodeNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    string BarcodeNo = txtBarcodeNo.Text.ToString();
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            BarcodeNo
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    oDS = clientMes2.ctEnumerateData("APISO.QryBarcodeProcessAttribute", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int nTotalMfgCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    JObject jResMfg = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = nTotalMfgCount,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });

                    string ItemValue = jResMfg["qrydata"][0]["ItemValue"].ToString();
                    BarcodeNo = jResMfg["qrydata"][0]["BarcodeNo"].ToString();
                    ItemValue = jResMfg["qrydata"][0]["ItemValue"].ToString();
                    if (Data=="Y") {
                        #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                        label = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                        label.PrintSettings.PrinterName = PrinterName;
                        label.Variables["Shift"].SetValue(ItemValue);
                        label.Variables["BarcodeNo"].SetValue(BarcodeNo);
                        label.Print(1);
                        #endregion
                    }
                    else {
                        #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                        label = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                        label.PrintSettings.PrinterName = PrinterName;
                        label.Variables["Shift"].SetValue("");
                        label.Variables["BarcodeNo"].SetValue(BarcodeNo);
                        label.Print(1);
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("異常:" + ex.ToString());
                }
            }
        }
    }
}
