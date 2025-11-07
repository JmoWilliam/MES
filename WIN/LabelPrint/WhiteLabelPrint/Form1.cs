using MESIII;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
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

namespace WhiteLabelPrint
{
    public partial class Form1 : Form
    {
        private ILabel label;
        public static ISDClient client = new ISDClient(0);
        string PrinterName = "", LabelPath = "";
        public static readonly IsoDateTimeConverter JsonConvertSetting = new IsoDateTimeConverter()
        {
            DateTimeFormat = "yyyy-MM-dd HH:mm:ss"
        };
        public Form1()
        {
            StreamReader srPrinterName = new StreamReader("D:\\WhiteLabelPrint\\Template\\PrinterName.txt", Encoding.Default);
            String linePrinterName;
            while ((linePrinterName = srPrinterName.ReadLine()) != null)
            {
                PrinterName = linePrinterName.ToString();
            }

            StreamReader srLabelPath = new StreamReader("D:\\WhiteLabelPrint\\Template\\LabelPath.txt", Encoding.Default);
            String lineLabelPath;
            while ((lineLabelPath = srLabelPath.ReadLine()) != null)
            {
                LabelPath = lineLabelPath.ToString();
            }
            InitializeComponent();
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

        private void TxMFG_ID_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    string BARCODE_NO1 = "", BARCODE_CAVITY1 = "";
                    string BARCODE_NO = "";
                    int MFG_ID = int.Parse(TxMFG_ID.Text.ToString());
                    #region //查詢ERP

                    #region //Request
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            MFG_ID,
                            BARCODE_NO
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = client.ctEnumerateData("MFGSO.QryMes2WhiteLabel001", jStr);
                    string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                    int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    #endregion
                    #region //Response
                    JObject jRes = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    #endregion

                    string MTL_NO = jRes["qrydata"][0]["MtlItemNo"].ToString();
                    string MTL_NAME = jRes["qrydata"][0]["MtlItemName"].ToString();
                    string MTL_NAME_EN = MTL_NAME;
                    #endregion
                    #region//條碼查詢

                    for (int i = 0; i < nTotalCount; i++)
                    {

                        BARCODE_NO1 = jRes["qrydata"][i]["BarcodeNo"].ToString();
                        label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                        label.PrintSettings.PrinterName = PrinterName;
                        label.Variables["BARCODE_NO"].SetValue(BARCODE_NO1);
                        label.Variables["BARCODE_CAVITY"].SetValue(BARCODE_NO1);
                        label.Variables["MTL_NAME_EN"].SetValue(MTL_NAME_EN);
                        label.Print(1);
                    }
                    clear();
                    #endregion

                }
            }
            catch (Exception e1)
            {
                string error_message = e1.Message;
                MESConfig.LogData(error_message);
            }
        }

        private void TxBarcodeNo_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    int MFG_ID = -1;
                    string BARCODE_NO = TxBarcodeNo.Text.ToString();
                    #region //查詢ERP
                    #region //Request
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            MFG_ID,
                            BARCODE_NO
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = client.ctEnumerateData("MFGSO.QryMes2WhiteLabel001", jStr);
                    string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                    #endregion
                    #region //Response
                    JObject jRes = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    #endregion
                    string MTL_NO = jRes["qrydata"][0]["MtlItemNo"].ToString();
                    string MTL_NAME = jRes["qrydata"][0]["MtlItemName"].ToString();
                    string MTL_NAME_EN = MTL_NAME;

                    #endregion
                    #region //Request
                    jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            MFG_ID,
                            BARCODE_NO
                        })
                    });
                    jStr = JsonConvert.SerializeObject(jObj);
                    oDS = client.ctEnumerateData("MFGSO.QryMes2WhiteLabel001", jStr);

                    #endregion

                    if (oDS.Tables[0].Rows.Count > 0)
                    {
                        label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                        label.PrintSettings.PrinterName = PrinterName;
                        label.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                        label.Variables["BARCODE_CAVITY"].SetValue(BARCODE_NO);
                        label.Variables["MTL_NAME_EN"].SetValue(MTL_NAME_EN);
                        label.Print(1);
                        clear();
                    }
                    else
                    {
                        MessageBox.Show("找不到該條碼");
                    }
                }
            }
            catch (Exception e1)
            {
                string error_message = e1.Message;
                MESConfig.LogData(error_message);
                MessageBox.Show(error_message);
            }
        }
        private void clear()
        {
            TxBarcodeNo.Text = "";
            TxMFG_ID.Text = "";
        }
    }
}
