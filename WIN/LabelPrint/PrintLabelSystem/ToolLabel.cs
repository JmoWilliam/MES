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
    public partial class ToolLabel : Form
    {
        private ILabel label;        
        int ToolCount = 0;
        string sData = "";
        public static Timer timer = new Timer();
        public ToolLabel()
        {
            InitializeComponent();
        }

        private void txtKnifeModel_ID_KeyDown(object sender, KeyEventArgs e)
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
                            ToolModelNo = txtKnifeModel_ID.Text.ToString(),
                            ToolNo =""
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryToolMode", jStr);
                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    ToolCount = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                    txtNotPrintLabelNum.Text = ToolCount.ToString();
                    #endregion
                }
            }
            catch (Exception ex)
            {
                MESConfig.LogData("工具標籤:" + ex.ToString());
                MessageBox.Show("工具標籤:" + ex.ToString());
            }
        }

        private void btnPrintLotLabel_Click(object sender, EventArgs e)
        {
            try
            {
                //欲列印數量不可以大於可列印數量
                if (int.Parse(txtNotPrintLabelNum.Text.ToString()) < int.Parse(txtPartNotPrintLabelNum.Text.ToString())) { throw new Exception("欲列印數量不可以大於可列印數量"); }

                //列印數量不可為0
                if (int.Parse(txtPartNotPrintLabelNum.Text.ToString())<=0) { throw new Exception("列印數量不可為0"); }

                //條碼種類不可以不等於 1或2
                if (txtBarcodeType.Text!="1" && txtBarcodeType.Text != "2") { throw new Exception("條碼類型 輸入錯誤"); }

                #region                
                string labelPath = "", printMachine = "";
                switch (txtBarcodeType.Text)
                {
                    case "1":
                        using (StreamReader labelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ToolLabel\\LabelName.txt", Encoding.Default))
                        {
                            labelPath = labelPathTxt.ReadToEnd();
                        }
                        break;
                    case "2":
                        using (StreamReader labelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ToolLabel\\LabelNameQR.txt", Encoding.Default))
                        {
                            labelPath = labelPathTxt.ReadToEnd();
                        }
                        break;
                }

                using (StreamReader printMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ToolLabel\\PrinterName.txt", Encoding.Default))
                {
                    printMachine = printMachineTxt.ReadToEnd();
                }

                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                label = PrintEngineFactory.PrintEngine.OpenLabel(labelPath);
                label.PrintSettings.PrinterName = printMachine;


                JObject TooljRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = ToolCount,
                    qrydata = JsonConvert.DeserializeObject(sData)
                }) ;
                int PartNotPrintLabelNum= int.Parse(txtPartNotPrintLabelNum.Text.ToString());
                for (int i = 0; i < PartNotPrintLabelNum; i++) {
                    timer.Interval = 1500;
                    timer.Start();
                    label.Variables["KNIFE_ID"].SetValue(TooljRes["qrydata"][i]["ToolNo"].ToString());
                    label.Variables["KNIFE_NO"].SetValue(TooljRes["qrydata"][i]["ToolModelNo"].ToString());
                    label.Variables["KNIFE_TYPE"].SetValue(TooljRes["qrydata"][i]["ToolCategoryName"].ToString());
                    label.Print(1);
                    timer.Stop();

                    #region //Request
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            ToolNo= TooljRes["qrydata"][i]["ToolNo"].ToString()
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    int nResult = -1;
                    nResult = MESConfig.clientNewMes.ctPostTxact("APISO.TxKnifePrint", jStr, TxTypeConsts.TxTypeUpdate);
                       
                    #endregion
                }
                TooljRes.RemoveAll();
                #endregion
            }
            catch (Exception ex) {
                MESConfig.LogData("工具標籤:" + ex.ToString());
                MessageBox.Show("工具標籤:" + ex.ToString());
            }
        }

        private void btnPrintLabel_Click(object sender, EventArgs e)
        {
            //條碼種類不可以不等於 1或2
            if (txttxtBarcodeTypePcs.Text != "1" && txttxtBarcodeTypePcs.Text != "2") { throw new Exception("條碼類型 輸入錯誤"); }

            #region
            string labelPath = "", printMachine = "";
            switch (txttxtBarcodeTypePcs.Text.ToString())
            {
                case "1":
                    using (StreamReader labelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ToolLabel\\LabelName.txt", Encoding.Default))
                    {
                        labelPath = labelPathTxt.ReadToEnd();
                    }
                    break;
                case "2":
                    using (StreamReader labelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ToolLabel\\LabelNameQR.txt", Encoding.Default))
                    {
                        labelPath = labelPathTxt.ReadToEnd();
                    }
                    break;
            }

            using (StreamReader printMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ToolLabel\\PrinterName.txt", Encoding.Default))
            {
                printMachine = printMachineTxt.ReadToEnd();
            }

            string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
            if (Directory.Exists(sdkFilesPath))
            {
                PrintEngineFactory.SDKFilesPath = sdkFilesPath;
            }
            PrintEngineFactory.PrintEngine.Initialize();
            label = PrintEngineFactory.PrintEngine.OpenLabel(labelPath);
            label.PrintSettings.PrinterName = printMachine;
            #endregion

            #region //查詢工具型號與明細
            JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            ToolModelNo = "",
                            ToolNo = txtToolNo.Text.ToString()
                        })
                    });
            string jStr = JsonConvert.SerializeObject(jObj);
            DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryToolMode", jStr);
            string ToolNosData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
            int Count = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

            JObject ToolNojRes = JObject.FromObject(new
            {
                status = "success",
                msg = "OK",
                table_rows = ToolCount,
                qrydata = JsonConvert.DeserializeObject(ToolNosData)
            }); ;
            timer.Interval = 1500;
            timer.Start();
            label.Variables["KNIFE_ID"].SetValue(ToolNojRes["qrydata"][0]["ToolNo"].ToString());
            label.Variables["KNIFE_NO"].SetValue(ToolNojRes["qrydata"][0]["ToolModelNo"].ToString());
            label.Variables["KNIFE_TYPE"].SetValue(ToolNojRes["qrydata"][0]["ToolCategoryName"].ToString());
            label.Print(1);
            timer.Stop();
            #region //Request
            jObj = JObject.FromObject(new
            {
                PARAMETER = JObject.FromObject(new
                {
                    ToolNo = ToolNojRes["qrydata"][0]["ToolNo"].ToString()
                })
            });
            jStr = JsonConvert.SerializeObject(jObj);
            int nResult = -1;
            nResult = MESConfig.clientNewMes.ctPostTxact("APISO.TxKnifePrint", jStr, TxTypeConsts.TxTypeUpdate);

            #endregion

            #endregion
        }
    }
}
