using MESIII;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NiceLabel.SDK;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PrintLabelSystem
{
    public partial class NewSunnyCustMoldLabel : Form
    {
        private Timer _timer;
        private ISDClient _clientMes2;
        private ISDClient _clientErp;
        private ILabel _label;
        private string _printerName;
        private string _labelName;

        private const string CONFIG_PATH = @"C:\WIN\PrintLabelSystem\Template\NewSunnyCustomerMoldLabel";

        private enum DatabaseType
        {
            MES,
            ERP
        }

        public NewSunnyCustMoldLabel()
        {
            InitializeComponent();
            InitializeTimer();
            InitializeDatabaseConnections();
            LoadConfiguration();
            InitializeControls();
        }

        private void InitializeTimer()
        {
            _timer = new Timer();
            _timer.Interval = 1500;
        }

        private void InitializeDatabaseConnections()
        {
            try
            {
                _clientMes2 = new ISDClient(0);
                _clientErp = new ISDClient(2);
            }
            catch (Exception ex)
            {
                HandleError("初始化資料庫連接失敗", ex);
                Application.Exit();
            }
        }

        private void LoadConfiguration()
        {
            try
            {
                _printerName = File.ReadAllText(Path.Combine(CONFIG_PATH, "PrinterName.txt"));
                _labelName = File.ReadAllText(Path.Combine(CONFIG_PATH, "LabelName.txt"));
            }
            catch (Exception ex)
            {
                HandleError("讀取配置文件失敗", ex);
                Application.Exit();
            }
        }

        private void InitializeControls()
        {
            cboStatus.Items.AddRange(new object[] { "新", "返", "镀" });
            cboStatus.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        protected override void OnCreateControl()
        {
            InitializePrintEngine();
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
                string sdkFilesPath = Path.Combine(
                    Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                    @"..\..\..\SDKFiles"
                );

                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
            }
            catch (SDKException ex)
            {
                HandleError("初始化 SDK 失敗", ex);
                Application.Exit();
            }
        }

        private void txtUseSandblasting_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            txtUseSandblasting.Text = txtUseSandblasting.Text.ToUpper() == "Y" ? "喷砂" : string.Empty;
        }

        private void txtMoldState_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            txtMoldState.Text = txtMoldState.Text.ToUpper() == "X" ? "X" : string.Empty;
        }

        private void txtMoId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;

            try
            {
                ValidateStatus();
                ProcessMoId();
            }
            catch (Exception ex)
            {
                HandleError("製令處理異常", ex);
            }
        }

        private void txtBarcodeNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;

            try
            {
                ValidateStatus();
                ProcessBarcode();
            }
            catch (Exception ex)
            {
                HandleError("條碼處理異常", ex);
            }
        }

        private void ValidateStatus()
        {
            if (string.IsNullOrEmpty(cboStatus.Text))
                throw new InvalidOperationException("模芯狀態不可不選");
        }

        private void ProcessMoId()
        {
            string mfg = txtMoId.Text;
            if (string.IsNullOrEmpty(mfg))
                throw new InvalidOperationException("製令不可空值");

            // 查詢工單資訊
            var orderInfo = QueryDatabase(
                "APISO.QrySunnyCustMoldLabel",
                new { WIP_TYPE = new { MFG = mfg } },
                DatabaseType.MES
            );
            if (orderInfo == null) return;

            // 查詢訂單資訊
            var soInfo = QueryDatabase(
                "APISO.QryErpSoDate",
                new
                {
                    PARAMETER = new
                    {
                        WoErpPrefix = orderInfo["WoErpPrefix"].ToString(),
                        WoErpNo = orderInfo["WoErpNo"].ToString()
                    }
                },
                DatabaseType.ERP
            );
            if (soInfo == null) return;

            // 查詢出貨資訊
            var deliveryInfo = GetDeliveryInfo(orderInfo, soInfo);
            if (deliveryInfo == null) return;

            PrintLabel(orderInfo, deliveryInfo);
        }

        private void ProcessBarcode()
        {
            string barcodeNo = txtBarcodeNo.Text;
            if (string.IsNullOrEmpty(barcodeNo))
                throw new InvalidOperationException("條碼不可空值");

            // 查詢工單資訊
            var wipInfo = QueryDatabase(
                "APISO.QryWipInfo",
                new
                {
                    PARAMETER = new
                    {
                        MFG = -1,
                        BarcodeNo = barcodeNo
                    }
                },
                DatabaseType.MES
            );
            if (wipInfo == null) return;

            // 查詢條碼資訊
            var orderInfo = QueryDatabase(
                "APISO.QryPcsSunnyCustMoldLabel",
                new
                {
                    WIP_TYPE = new
                    {
                        BARCODE = barcodeNo
                    }
                },
                DatabaseType.MES
            );
            if (orderInfo == null) return;

            // 查詢訂單和出貨資訊
            var soInfo = GetSalesOrderInfo(wipInfo);
            if (soInfo == null) return;

            var deliveryInfo = GetDeliveryInfo(wipInfo, soInfo);
            if (deliveryInfo == null) return;

            PrintLabel(orderInfo, deliveryInfo);
        }

        private JToken QueryDatabase(string apiName, object queryParams, DatabaseType dbType)
        {
            try
            {
                string jsonStr = JsonConvert.SerializeObject(queryParams);
                DataSet oDS;

                switch (dbType)
                {
                    case DatabaseType.MES:
                        oDS = _clientMes2.ctEnumerateData(apiName, jsonStr);
                        break;
                    case DatabaseType.ERP:
                        oDS = _clientErp.ctEnumerateData(apiName, jsonStr);
                        break;
                    default:
                        throw new ArgumentException("無效的資料庫類型");
                }

                if (!ValidateDataSet(oDS, apiName))
                    return null;

                string tableData = JsonConvert.SerializeObject(
                    oDS.Tables[0],
                    Formatting.Indented,
                    MESConfig.JsonConvertSetting
                );

                var result = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]),
                    qrydata = JsonConvert.DeserializeObject(tableData)
                });

                return result["qrydata"][0];
            }
            catch (Exception ex)
            {
                HandleError(string.Format("查詢 {0} 失敗", apiName), ex);
                return null;
            }
        }

        private bool ValidateDataSet(DataSet ds, string apiName)
        {
            if (ds?.Tables == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
            {
                MessageBox.Show(string.Format("查詢 {0} 無資料", apiName));
                return false;
            }
            return true;
        }

        private JToken GetSalesOrderInfo(JToken wipInfo)
        {
            return QueryDatabase(
                "APISO.QryErpSoDate",
                new
                {
                    PARAMETER = new
                    {
                        WoErpPrefix = wipInfo["WoErpPrefix"].ToString(),
                        WoErpNo = wipInfo["WoErpNo"].ToString()
                    }
                },
                DatabaseType.ERP
            );
        }

        private JToken GetDeliveryInfo(JToken orderInfo, JToken soInfo)
        {
            return QueryDatabase(
                "APISO.QryDeliverySchedule",
                new
                {
                    PARAMETER = new
                    {
                        MtlItemNo = orderInfo["MtlItemNo"].ToString(),
                        SoErpPrefix = soInfo["SoErpPrefix"].ToString(),
                        SoErpNo = soInfo["SoErpNo"].ToString(),
                        SoSequence = soInfo["SoSequence"].ToString()
                    }
                },
                DatabaseType.MES
            );
        }

        private void PrintLabel(JToken orderInfo, JToken deliveryInfo)
        {
            _timer.Start();

            try
            {
                _label = PrintEngineFactory.PrintEngine.OpenLabel(_labelName);
                _label.PrintSettings.PrinterName = _printerName;

                SetLabelVariables(orderInfo, deliveryInfo);
                _label.Print(1);
            }
            catch (Exception ex)
            {
                HandleError("列印標籤失敗", ex);
            }
            finally
            {
                _timer.Stop();
                if (_label != null)
                {
                    _label.Dispose();
                    _label = null;
                }
            }
        }

        private void SetLabelVariables(JToken orderInfo, JToken deliveryInfo)
        {
            _label.Variables["WIP_NAME"].SetValue(orderInfo["BarcodeItem"].ToString());
            _label.Variables["BARCODE_NO"].SetValue(orderInfo["BarcodeNo"].ToString());
            _label.Variables["ShipDate"].SetValue(deliveryInfo["PcPromiseDate"].ToString());
            _label.Variables["Company"].SetValue(GetCompanyName(deliveryInfo["Column1"].ToString()));
            _label.Variables["UseSandblasting"].SetValue(txtUseSandblasting.Text);
            _label.Variables["moldState"].SetValue(txtMoldState.Text);
            _label.Variables["Status"].SetValue(cboStatus.Text);
        }

        private string GetCompanyName(string originalName)
        {
            switch (originalName)
            {
                case "中揚光電":
                    return "中扬";
                case "晶彩光学":
                    return "晶彩";
                case "紘立光電":
                    return "紘立";
                default:
                    return "中扬";
            }
        }

        private void HandleError(string message, Exception ex)
        {
            MessageBox.Show(string.Format("{0}: {1}", message, ex.Message));
        }


    }
}