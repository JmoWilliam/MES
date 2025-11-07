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
    public partial class SunnyJCShip : Form
    {
        public static Timer timer = new Timer();
        public static ISDClient clientMes2 = new ISDClient(0);
        public static ISDClient clientErp = new ISDClient(2);
        private ILabel label, labe2;
        string UseSandblasting = "", sData = "";
        string PrinterName = "", LabelName = "";
        string ErpNo="", MtlItemNo = "",MtlItemName="", MtlItemSpec="", PlanQty="",BarcodeNo="";
        string SoErpPrefix = "", SoErpNo = "", SoSequence = "", CompanyName = "";
        string PcPromiseDate = "";
        int Moid = -1;
        string MoidText = "";

        int PAGE_INDEX = -1, PAGE_SIZE = -1;
        string ORDER_BY = "";        
        DataSet oDS;
        public SunnyJCShip()
        {
            InitializeComponent();

            StreamReader txLabelNamesr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyJCShip\\LabelName.txt", Encoding.Default);
            String lineLabelName;
            while ((lineLabelName = txLabelNamesr.ReadLine()) != null)
            {
                LabelName = lineLabelName.ToString();
            }

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
            newRow["LabelPrintNo"] = ""; // 假设 -1 代表无效值或默认值
            dt.Rows.InsertAt(newRow, 0);
            cboLabelMachine.DisplayMember = "LabelPrintName";
            cboLabelMachine.ValueMember = "LabelPrintNo";
            cboLabelMachine.DataSource = dt;
            cboLabelMachine.SelectedIndex = 0;

            //https://chatgpt.com/c/f945f423-b5bc-4c2b-8518-dd2c516ea083
            #endregion
            cboAtomizationMethod.SelectedIndex = 0;
            cboStatus.SelectedIndex = 0;
            groupBoxMo.Visible = false;
            groupBoxBarcode.Visible = false;
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
        private void txtMoId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try {
                    txtBarcodeNo.Text = "";
                    MoidText = txtMoId.Text.ToString();
                    if (MoidText == "")
                    {
                        throw new Exception("製令不可空值");
                    }
                    else {
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
                    oDS = clientMes2.ctEnumerateData("APISO.QryManufactureOrder", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int nTotalMfgCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                    JObject jResMfg = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = nTotalMfgCount,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    ErpNo = jResMfg["qrydata"][0]["ErpNo"].ToString();
                    labErpWoNo.Text= ErpNo;

                    MtlItemNo = jResMfg["qrydata"][0]["MtlItemNo"].ToString();
                    labMtlItemNo.Text = MtlItemNo;
                    MtlItemName = jResMfg["qrydata"][0]["MtlItemName"].ToString();
                    labMtlItemName.Text = MtlItemName;
                    MtlItemSpec = jResMfg["qrydata"][0]["MtlItemSpec"].ToString();
                    labMtlItemSpec.Text = MtlItemSpec;
                    PlanQty = jResMfg["qrydata"][0]["PlanQty"].ToString();
                    labPlanQty.Text = PlanQty;

                    groupBoxMo.Visible = true;
                    groupBoxBarcode.Visible = false;


                }
                catch (Exception ex)
                {
                    MessageBox.Show("製令異常:" + ex.ToString());
                }
            }
        }
        private void txtBarcodeNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    txtMoId.Text = "";
                    string BarcodeNoText = txtBarcodeNo.Text.ToString();
                    if (BarcodeNoText == "")
                    {
                        throw new Exception("條碼不可空值");
                    }
                    else
                    {
                        BarcodeNo = BarcodeNoText;
                    }
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            BarcodeNo
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    oDS = clientMes2.ctEnumerateData("APISO.QryManufactureOrderBarcodeNo", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int nTotalMfgCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                    JObject jResMfg = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = nTotalMfgCount,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    ErpNo = jResMfg["qrydata"][0]["ErpNo"].ToString();
                    labErpWoNoB.Text = ErpNo;
                    MtlItemNo = jResMfg["qrydata"][0]["MtlItemNo"].ToString();
                    labMtlItemNoB.Text = MtlItemNo;
                    MtlItemName = jResMfg["qrydata"][0]["MtlItemName"].ToString();
                    labMtlItemNameB.Text = MtlItemName;
                    MtlItemSpec = jResMfg["qrydata"][0]["MtlItemSpec"].ToString();
                    labMtlItemSpecB.Text = MtlItemSpec;
                    PlanQty = jResMfg["qrydata"][0]["PlanQty"].ToString();
                    labPlanQtyB.Text = PlanQty;
                    string ItemValue = jResMfg["qrydata"][0]["ItemValue"].ToString();
                    label11ItemValue.Text = ItemValue;


                    groupBoxMo.Visible = false; 
                    groupBoxBarcode.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("條碼異常:" + ex.ToString());
                }
            }
        }
        private void txtMoItemPartId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    txtMoId.Text = "";
                    string MoItemPartId = txtMoItemPartId.Text.ToString();
                    if (MoItemPartId == "")
                    {
                        throw new Exception("刻字條碼不可空值");
                    }
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            MoItemPartId
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    oDS = clientMes2.ctEnumerateData("APISO.QryManufactureOrderMoItemPart", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int nTotalMfgCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                    JObject jResMfg = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = nTotalMfgCount,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    txtBarcodeNo.Text= jResMfg["qrydata"][0]["BarcodeNo"].ToString();
                    ErpNo = jResMfg["qrydata"][0]["ErpNo"].ToString();
                    labErpWoNoB.Text = ErpNo;
                    MtlItemNo = jResMfg["qrydata"][0]["MtlItemNo"].ToString();
                    labMtlItemNoB.Text = MtlItemNo;
                    MtlItemName = jResMfg["qrydata"][0]["MtlItemName"].ToString();
                    labMtlItemNameB.Text = MtlItemName;
                    MtlItemSpec = jResMfg["qrydata"][0]["MtlItemSpec"].ToString();
                    labMtlItemSpecB.Text = MtlItemSpec;
                    PlanQty = jResMfg["qrydata"][0]["PlanQty"].ToString();
                    labPlanQtyB.Text = PlanQty;

                    groupBoxMo.Visible = false;
                    groupBoxBarcode.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("條碼異常:" + ex.ToString());
                }
            }
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrinterName = cboLabelMachine.SelectedValue.ToString();
            if (PrinterName != "")
            {
                PrinterName = cboLabelMachine.SelectedValue.ToString();
            }
            else {
                StreamReader sr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyJCShip\\PrinterName.txt", Encoding.Default);
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    PrinterName = line.ToString();
                }
            }
            UseSandblasting = cboAtomizationMethod.Text.ToString();
            if (UseSandblasting=="无") {
                UseSandblasting = "";
            }

            string Status = cboStatus.Text.ToString();
            if (Status == "无")
            {
                Status = "";
            }
            //訂單的排定交貨日
            DateTime selectedStartDate = mCalendarPackingDate.SelectionStart;
            PcPromiseDate = selectedStartDate.ToString("yyyy/MM/dd");                   
            if (txtMoId.Text != "" && txtBarcodeNo.Text == "") {
                try
                {
                    string MFG = txtMoId.Text.ToString();
                    if (MFG == "") throw new Exception("製令不可空值");
                    
                    JObject jObj = JObject.FromObject(new
                    {
                        WIP_TYPE = JObject.FromObject(new
                        {
                            MFG
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    oDS = clientMes2.ctEnumerateData("APISO.QrySunnyCustMoldLabel", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int nTotalMfgCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    JObject jResMfg = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = nTotalMfgCount,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    if (nTotalMfgCount > 0)
                    {
                        for (int i = 0; i < nTotalMfgCount; i++)
                        {
                            timer.Interval = 1500;
                            timer.Start();
                            string BarcodeItem = jResMfg["qrydata"][i]["BarcodeItem"].ToString();
                            string lastNo = BarcodeItem.Split('-').Last();                            
                            string BARCODE_NO = jResMfg["qrydata"][i]["BarcodeNo"].ToString();

                            #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                            labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                            labe2.PrintSettings.PrinterName = PrinterName;
                            labe2.Variables["WIP_NAME"].SetValue(BarcodeItem);
                            labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                            labe2.Variables["ShipDate"].SetValue(PcPromiseDate);
                            labe2.Variables["Company"].SetValue("晶彩");
                            labe2.Variables["UseSandblasting"].SetValue(UseSandblasting);
                            labe2.Variables["Status"].SetValue(Status);
                            labe2.Print(1);
                            #endregion

                            timer.Stop();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("製令列印異常:" + ex.ToString());
                }
            } else if (txtMoId.Text == "" && txtBarcodeNo.Text != "") {
                try
                {
                    string BarcodeNo = txtBarcodeNo.Text.ToString();
                    if (BarcodeNo == "") throw new Exception("條碼不可空值");
                    JObject jObj = JObject.FromObject(new
                    {
                        WIP_TYPE = JObject.FromObject(new
                        {
                            BARCODE = txtBarcodeNo.Text.ToString()
                        }),
                        PAGE_INFO = JObject.FromObject(new
                        {
                            PAGE_INDEX,
                            PAGE_SIZE,
                            ORDER_BY
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    oDS = clientMes2.ctEnumerateData("APISO.QryPcsSunnyCustMoldLabelJc", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    JObject jRes = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = nTotalCount,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });

                    timer.Interval = 1500;
                    timer.Start();
                    string BarcodeItem = jRes["qrydata"][0]["BarcodeItem"].ToString();
                    string lastNo = BarcodeItem.Split('-').Last();                    
                    string BARCODE_NO = jRes["qrydata"][0]["BarcodeNo"].ToString();

                    #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                    labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                    labe2.PrintSettings.PrinterName = PrinterName;
                    labe2.Variables["WIP_NAME"].SetValue(BarcodeItem);
                    labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                    labe2.Variables["ShipDate"].SetValue(PcPromiseDate);
                    labe2.Variables["Company"].SetValue("晶彩");
                    labe2.Variables["UseSandblasting"].SetValue(UseSandblasting);
                    labe2.Variables["Status"].SetValue(Status);
                    labe2.Print(1);
                    #endregion

                    timer.Stop();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("條碼列印異常:" + ex.ToString());
                }
            }
        }


    }
}
