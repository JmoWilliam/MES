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

namespace PrintLabelSystem
{
    public partial class SunnyCustMoldLabel : Form
    {
        public static Timer timer = new Timer();       
        public static ISDClient clientMes2 = new ISDClient(0);
        public static ISDClient clientErp = new ISDClient(2);
        private ILabel label, labe2;
        string PRINT_PERMISSION = "Y"; //預設列印權限 Y:可以全部印;N:只能印部分
        int USER_ID = -1;
        int PAGE_INDEX = -1, PAGE_SIZE = -1;
        string ORDER_BY = "";
        string BARCODE_NO = "";
        string Password = "";
        string USER_NO = "";
        string PrinterName = "", LabelName = "";
        DataSet oDS;
        public static CmnSrvLib cmn = new CmnSrvLib();
        public static readonly IsoDateTimeConverter JsonConvertSetting = new IsoDateTimeConverter()
        {
            DateTimeFormat = "yyyy-MM-dd HH:mm:ss"
        };
        public SunnyCustMoldLabel()
        {
            InitializeComponent();
            StreamReader sr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyCustomerMoldLabel\\PrinterName.txt", Encoding.Default);
            //StreamReader sr = new StreamReader("D:\\JMO Projects\\JMO_Prod_New\\WIN\\LabelPrint\\PrintLabelSystem\\Template\\SunnyCustomerMoldLabel\\PrinterName.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                PrinterName = line.ToString();
            }
            StreamReader txLabelNamesr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyCustomerMoldLabel\\LabelName.txt", Encoding.Default);
            //StreamReader txLabelNamesr = new StreamReader("D:\\JMO Projects\\JMO_Prod_New\\WIN\\LabelPrint\\PrintLabelSystem\\Template\\SunnyCustomerMoldLabel\\LabelName.txt", Encoding.Default);
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

        private void txtCheckNi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) {

            }
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

        private void txMFG_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txMFG.Text.ToString() == "")
                {
                    MessageBox.Show("★製令條碼不能為空!!");
                }
                else
                {
                    if (txMFG.Text.ToString() != "")
                    {
                        string MTL_NAME = "";
                        int BARCODE_ID = -1;                        
                        string WoErpNo = "", WoErpPrefix = "", MtlItemNo = "";
                        string SoErpPrefix = "", SoErpNo = "", CompanyName ="";
                        string PcPromiseDate = "";
                        int MFG = int.Parse(txMFG.Text.ToString());

                        #region //取得 工單 單號單別
                        JObject jObjWo = JObject.FromObject(new
                        {
                            PARAMETER = JObject.FromObject(new
                            {
                                MFG,
                                BarcodeNo=""
                            })
                        });
                        string jStrWo = JsonConvert.SerializeObject(jObjWo);
                        oDS = clientMes2.ctEnumerateData("APISO.QryWipInfo", jStrWo);
                        string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                        int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                        JObject jRes = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "OK",
                            table_rows = TOTAL_COUNT,
                            qrydata = JsonConvert.DeserializeObject(sData)
                        });
                        WoErpNo = jRes["qrydata"][0]["WoErpNo"].ToString();
                        WoErpPrefix = jRes["qrydata"][0]["WoErpPrefix"].ToString();
                        MtlItemNo = jRes["qrydata"][0]["MtlItemNo"].ToString();
                        //查詢訂單 單號單別
                        //JObject jObjErp = JObject.FromObject(new
                        //{
                        //    PARAMETER = JObject.FromObject(new
                        //    {
                        //        WoErpPrefix,
                        //        WoErpNo
                        //    })
                        //});
                        //string jStrErp = JsonConvert.SerializeObject(jObjErp);
                        //oDS = clientErp.ctEnumerateData("APISO.QryErpSoDate", jStrErp);
                        //sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                        //TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                        //jRes = JObject.FromObject(new
                        //{
                        //    status = "success",
                        //    msg = "OK",
                        //    table_rows = TOTAL_COUNT,
                        //    qrydata = JsonConvert.DeserializeObject(sData)
                        //});
                        //SoErpPrefix = jRes["qrydata"][0]["SoErpPrefix"].ToString();
                        //SoErpNo = jRes["qrydata"][0]["SoErpNo"].ToString();

                        //查詢訂單的 排定交貨日
                        //JObject jObjSo = JObject.FromObject(new
                        //{
                        //    PARAMETER = JObject.FromObject(new
                        //    {
                        //        MtlItemNo
                        //    })
                        //});
                        //string jStrSo = JsonConvert.SerializeObject(jObjSo);
                        //oDS = clientMes2.ctEnumerateData("APISO.QryDeliverySchedule", jStrSo);
                        //sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                        //TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                        //jRes = JObject.FromObject(new
                        //{
                        //    status = "success",
                        //    msg = "OK",
                        //    table_rows = TOTAL_COUNT,
                        //    qrydata = JsonConvert.DeserializeObject(sData)
                        //});
                        //PcPromiseDate = jRes["qrydata"][0]["PcPromiseDate"].ToString();
                        //switch (jRes["qrydata"][0]["CompanyName"].ToString()) {
                        //    case "中揚光電":
                        //        CompanyName = "中揚";
                        //        break;
                        //    case "晶彩光学":
                        //        CompanyName = "晶彩";
                        //        break;
                        //    case "紘立光電":
                        //        CompanyName = "紘立";
                        //        break;
                        //    default:
                        //        CompanyName = "中揚";
                        //        break;
                        //}                        
                        #endregion

                        #region //Request
                        JObject jObj = JObject.FromObject(new
                        {
                            WIP_TYPE = JObject.FromObject(new
                            {
                                MFG,
                                PRINT_PERMISSION
                            }),
                            PAGE_INFO = JObject.FromObject(new
                            {
                                PAGE_INDEX,
                                PAGE_SIZE,
                                ORDER_BY
                            })
                        });
                        string jStr = JsonConvert.SerializeObject(jObj);
                        if (PRINT_PERMISSION == "Y")
                        {
                            oDS = clientMes2.ctEnumerateData("APISO.QrySunnyCustMoldLabel", jStr);
                        }
                        else
                        {
                            oDS = clientMes2.ctEnumerateData("APISO.QrySunnyCustMoldLabel", jStr);
                        }
                        sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                        int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                        #endregion

                        #region //Response
                        jRes = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "OK",
                            table_rows = nTotalCount,
                            qrydata = JsonConvert.DeserializeObject(sData)
                        });
                        #endregion

                        if (nTotalCount == 0)
                        {
                            MessageBox.Show("☆此製令未綁定模仁條碼!!");

                        }
                        else
                        {
                            for (int i = 0; i < nTotalCount; i++)
                            {
                                timer.Interval = 1000;
                                timer.Start();
                                string BarcodeItem = jRes["qrydata"][i]["BarcodeItem"].ToString();
                                string lastNo = BarcodeItem.Split('-').Last();
                                string MtlItemName = jRes["qrydata"][i]["MtlItemName"].ToString().Split(' ')[1] + '-' + lastNo;
                                BARCODE_NO = jRes["qrydata"][i]["BarcodeNo"].ToString();
                                string Remark = "";
                                if (txtCheckNi.Text.ToString()== "OK") {
                                    Remark = "退鎳重鍍";
                                }

                                #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                                labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                                labe2.PrintSettings.PrinterName = PrinterName;
                                labe2.Variables["WIP_NAME"].SetValue(BarcodeItem);
                                labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                                labe2.Variables["ShipDate"].SetValue(PcPromiseDate);
                                labe2.Variables["Company"].SetValue(CompanyName);
                                labe2.Variables["Remark"].SetValue(Remark);
                                labe2.Print(1);
                                #endregion

                                timer.Stop();
                            }
                        }
                    }
                }
                txMFG.Text = "";//清空
            }
        }

        private void TxBarCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (TxBarCode.Text.ToString() != "")
                {
                    if (TxBarCode.Text.ToString() == "")
                    {
                        MessageBox.Show("★模仁條碼不能為空!!");
                    }
                    else
                    {
                        if (TxBarCode.Text.ToString() != "")
                        {
                            string BARCODE = TxBarCode.Text.ToString();
                            string MFG = "";
                            string WoErpNo = "", WoErpPrefix = "", MtlItemNo = "";
                            string SoErpPrefix = "", SoErpNo = "", CompanyName = "";
                            string PcPromiseDate = "";

                            #region //取得 工單 單號單別
                            JObject jObjWo = JObject.FromObject(new
                            {
                                PARAMETER = JObject.FromObject(new
                                {
                                    MFG=-1,
                                    BarcodeNo = BARCODE
                                })
                            });
                            string jStrWo = JsonConvert.SerializeObject(jObjWo);
                            DataSet oDS = clientMes2.ctEnumerateData("APISO.QryWipInfo", jStrWo);
                            string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                            int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                            JObject jRes = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "OK",
                                table_rows = TOTAL_COUNT,
                                qrydata = JsonConvert.DeserializeObject(sData)
                            });
                            WoErpNo = jRes["qrydata"][0]["WoErpNo"].ToString();
                            WoErpPrefix = jRes["qrydata"][0]["WoErpPrefix"].ToString();
                            MtlItemNo = jRes["qrydata"][0]["MtlItemNo"].ToString();                            

                            //查詢訂單的 排定交貨日
                            //JObject jObjSo = JObject.FromObject(new
                            //{
                            //    PARAMETER = JObject.FromObject(new
                            //    {
                            //        MtlItemNo
                            //    })
                            //});
                            //string jStrSo = JsonConvert.SerializeObject(jObjSo);
                            //oDS = clientMes2.ctEnumerateData("APISO.QryDeliverySchedule", jStrSo);
                            //sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                            //TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                            //jRes = JObject.FromObject(new
                            //{
                            //    status = "success",
                            //    msg = "OK",
                            //    table_rows = TOTAL_COUNT,
                            //    qrydata = JsonConvert.DeserializeObject(sData)
                            //});
                            //PcPromiseDate = jRes["qrydata"][0]["PcPromiseDate"].ToString();
                            //switch (jRes["qrydata"][0]["CompanyName"].ToString())
                            //{
                            //    case "中揚光電":
                            //        CompanyName = "中揚";
                            //        break;
                            //    case "晶彩光学":
                            //        CompanyName = "晶彩";
                            //        break;
                            //    case "紘立光電":
                            //        CompanyName = "紘立";
                            //        break;
                            //    default:
                            //        CompanyName = "中揚";
                            //        break;
                            //}

                            #endregion

                            #region //Request
                            JObject jObj = JObject.FromObject(new
                            {
                                WIP_TYPE = JObject.FromObject(new
                                {
                                    BARCODE
                                }),
                                PAGE_INFO = JObject.FromObject(new
                                {
                                    PAGE_INDEX,
                                    PAGE_SIZE,
                                    ORDER_BY
                                })
                            });
                            string jStr = JsonConvert.SerializeObject(jObj);
                            oDS = clientMes2.ctEnumerateData("APISO.QryPcsSunnyCustMoldLabel", jStr);
                            sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                            int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                            #endregion

                            #region //Response
                            jRes = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "OK",
                                table_rows = nTotalCount,
                                qrydata = JsonConvert.DeserializeObject(sData)
                            });
                            #endregion 

                            string BarcodeItem = jRes["qrydata"][0]["BarcodeItem"].ToString();
                            string lastNo = BarcodeItem.Split('-').Last();
                            string MtlItemName = jRes["qrydata"][0]["MtlItemName"].ToString().Split(' ')[1] + '-' + lastNo;
                            BARCODE_NO = jRes["qrydata"][0]["BarcodeNo"].ToString();
                            string Remark = "";
                            if (txtCheckNi.Text.ToString() == "OK")
                            {
                                Remark = "退鎳重鍍";
                            }

                            #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                            labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                            labe2.PrintSettings.PrinterName = PrinterName;

                            labe2.Variables["WIP_NAME"].SetValue(BarcodeItem);
                            labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                            labe2.Variables["ShipDate"].SetValue(PcPromiseDate);
                            labe2.Variables["Company"].SetValue(CompanyName);
                            labe2.Variables["Remark"].SetValue(Remark);
                            labe2.Print(1);
                            #endregion
                            TxBarCode.Text = "";//清空
                        }
                    }
                }
                else
                {
                    MessageBox.Show("請輸入產品條碼");
                    TxBarCode.Text = "";//清空                     
                }
            }
        }
    }
}
