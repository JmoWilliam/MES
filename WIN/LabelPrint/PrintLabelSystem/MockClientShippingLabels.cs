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
    public partial class MockClientShippingLabels : Form
    {
        public string Data { get; set; }
        public static ISDClient clientMes2 = new ISDClient(0);
        public static ISDClient clientErp = new ISDClient(2);
        private ILabel label;
        public static Timer timer = new Timer();
        string customerName; //客戶名稱
        string customerPurchaseOrder; //客戶PO
        string customerDepartmentNumber; //客戶部番
        string mtlItemName;//廠內部番
        string productBarcode;//廠內Lot.No (產品條碼+班次)
        string barcodeNo;//產品條碼
        int barcodeQty;//數量
        string packerUserNo;//包裝員工號
        string packingDate;//包裝日期(當日列印日期
        string remark;//備註
        string PrinterName = "", LabelName = "";
        public MockClientShippingLabels()
        {
            InitializeComponent();

            StreamReader sr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\MockClientShippingLabels\\PrinterName.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                PrinterName = line.ToString();
            }
            StreamReader txLabelNamesr = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\MockClientShippingLabels\\LabelName.txt", Encoding.Default);
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
        private void btnPrint_Click(object sender, EventArgs e)
        {
            barcodeNo = txtBarcodeNo.Text.ToString();
            packerUserNo= txtPackerUserNo.Text.ToString();
            if (!barcodeNo.Equals("") || !packerUserNo.Equals(""))
            {
                try {
                    JObject jObjWo = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            ItemNo= "Shift",
                            barcodeNo 
                        })
                    });
                    string jStrWo = JsonConvert.SerializeObject(jObjWo);
                    DataSet oDS = clientMes2.ctEnumerateData("APISO.QryBarcodeAttribute02", jStrWo);
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
                    mtlItemName = jRes["qrydata"][0]["MtlItemName"].ToString();
                    barcodeQty= int.Parse(jRes["qrydata"][0]["BarcodeQty"].ToString());
                    string ItemValue = jRes["qrydata"][0]["ItemValue"].ToString();

                    DateTime selectedStartDate = mCalendarPackingDate.SelectionStart;
                    packingDate = selectedStartDate.ToString("yyyy/MM/dd");

                    if (!MTL_NO.Equals(""))
                    {
                        JObject jObj = JObject.FromObject(new
                        {
                            PARAMETER = JObject.FromObject(new
                            {
                                MTL_NO
                            })
                        });
                        string jStrErp = JsonConvert.SerializeObject(jObj);
                        oDS = MESConfig.clientErp.ctEnumerateData("APISO.QryCustomerItem02", jStrErp);
                        sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                        TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                        JObject jResErp = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "OK",
                            table_rows = TOTAL_COUNT,
                            qrydata = JsonConvert.DeserializeObject(sData)
                        });
                        customerName = jResErp["qrydata"][0]["MA002"].ToString();
                        customerDepartmentNumber = jResErp["qrydata"][0]["MG005"].ToString();

                        #region//列印標籤
                        if (Data=="Y") {
                            timer.Interval = 1500;
                            timer.Start();
                            label = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                            label.PrintSettings.PrinterName = PrinterName;                           
                            label.Variables["barcodeQty"].SetValue(barcodeQty.ToString());
                            label.Variables["customerDepartmentNumber"].SetValue(customerDepartmentNumber);
                            label.Variables["customerName"].SetValue(customerName);
                            label.Variables["customerPurchaseOrder"].SetValue(customerPurchaseOrder);
                            label.Variables["mtlItemName"].SetValue(mtlItemName);
                            label.Variables["packerUserNo"].SetValue(packerUserNo);
                            label.Variables["packingDate"].SetValue(packingDate);
                            label.Variables["productBarcode"].SetValue(barcodeNo);
                            label.Variables["ItemValue"].SetValue(ItemValue);
                            label.Variables["remark"].SetValue(txtRemark.Text.ToString());
                            label.Print(1);
                            timer.Stop();
                        }
                        else {  
                            timer.Interval = 1500;
                            timer.Start();
                            label = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                            label.PrintSettings.PrinterName = PrinterName;                            
                            label.Variables["barcodeQty"].SetValue(barcodeQty.ToString());
                            label.Variables["customerDepartmentNumber"].SetValue(customerDepartmentNumber);
                            label.Variables["customerName"].SetValue(customerName);
                            label.Variables["customerPurchaseOrder"].SetValue(customerPurchaseOrder);
                            label.Variables["mtlItemName"].SetValue(mtlItemName);
                            label.Variables["packerUserNo"].SetValue(packerUserNo);
                            label.Variables["packingDate"].SetValue(packingDate);
                            label.Variables["productBarcode"].SetValue(barcodeNo);
                            label.Variables["ItemValue"].SetValue("");
                            label.Variables["remark"].SetValue(txtRemark.Text.ToString());
                            label.Print(1);
                            timer.Stop();
                        }
                        #endregion
                    }
                    else {
                        throw new MyException("查無品號");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("異常:" + ex.ToString());
                }
            }
            else {
                MessageBox.Show("【產品條碼】或【包裝員工號】不可輸入空值");
            }
        }
    }
}
