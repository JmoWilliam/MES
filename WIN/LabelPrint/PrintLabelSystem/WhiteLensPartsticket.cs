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
    public partial class WhiteLensPartsticket : Form
    {
        private ILabel label;
        public WhiteLensPartsticket()
        {
            InitializeComponent();

        }

        private void txtBarcodeNo_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    #region //成形部品票資訊
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            MoId = -1,
                            BARCODE_NO = txtBarcodeNo.Text.ToString()
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryLensPartsticketDate", jStr);
                    string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                    int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    JObject jRes = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        table_rows = TOTAL_COUNT,
                        qrydata = JsonConvert.DeserializeObject(sData)
                    });
                    #endregion

                    var finishDateTime = DateTime.Parse(jRes["qrydata"][0]["FinishDate"].ToString());                    
                    var hour = finishDateTime.Hour;

                    var TIME_NO = "";
                    if (hour >= 7 && hour < 11)
                    {
                        TIME_NO = "A(7~11點)";
                    }
                    else if (hour >= 11 && hour < 15)
                    {
                        TIME_NO = "B(11~15點)";
                    }
                    else if (hour >= 15 && hour < 19)
                    {
                        TIME_NO = "C(15~19點)";
                    }
                    else if (hour >= 19 && hour < 23)
                    {
                        TIME_NO = "E(19~23點)";
                    }
                    else if (hour >= 23 || hour < 3)
                    {
                        TIME_NO = "F(23~3點)";
                    }
                    else if (hour >= 3 && hour < 7)
                    {
                        TIME_NO = "G(3~7點)";
                    }

                    txtCustItemNo.Text = jRes["qrydata"][0]["MtlItemNo"].ToString();//部番
                    txtCavityValue.Text = jRes["qrydata"][0]["ItemValue"].ToString();//穴號
                    txtBarcodeQty.Text= jRes["qrydata"][0]["BarcodeQty"].ToString();//數量                       
                    string []FinishDateArray = jRes["qrydata"][0]["FinishDate"].ToString().Split(' ');
                    txtFinishDate.Text = FinishDateArray[0].ToString();//日期
                    txtFinishTime.Text = TIME_NO;//時間
                    txtUserNo.Text = jRes["qrydata"][0]["UserNo"].ToString();

                    int BarcodeProcessId = int.Parse(jRes["qrydata"][0]["BarcodeProcessId"].ToString());
                    string BarcodeNo = jRes["qrydata"][0]["BarcodeNo"].ToString();//產品條碼

                    #region //修改穴號
                    //int CavityValue = int.Parse(txtCavityValue.Text.ToString());

                    //jObj = JObject.FromObject(new
                    //{
                    //    PARAMETER = JObject.FromObject(new
                    //    {
                    //        BarcodeNo,
                    //        CavityValue
                    //    })
                    //});
                    //jStr = JsonConvert.SerializeObject(jObj);

                    //int nResult = -1;
                    //nResult = MESConfig.clientNewMes.ctPostTxact("APISO.TxCavityValue", jStr, TxTypeConsts.TxTypeUpdate);


                    //jRes = JObject.FromObject(new
                    //{
                    //    status = "success",
                    //    msg = "OK",
                    //    qrydata = nResult
                    //});
                    #endregion
                }
                
            }
            catch (Exception ex)
            {
                MESConfig.LogData("白物成型部品票標籤:" + ex.ToString());
            }
        }

        private void btnPrintUpload_Click(object sender, EventArgs e)
        {
            //#region //標籤樣式設定+印表機選擇                            
            //string LabelPath = "", LabelPathLine;
            //string PrintMachine = "", PrintMachineLine;

            //StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\WhiteLensPartsticket\\LabelName.txt", Encoding.Default);
            //while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
            //{
            //    LabelPath = LabelPathLine.ToString();
            //}

            //StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\WhiteLensPartsticket\\PrinterName.txt", Encoding.Default);
            //while ((PrintMachineLine = PrintMachineTxt.ReadLine()) != null)
            //{
            //    PrintMachine = PrintMachineLine.ToString();
            //}

            //string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
            //if (Directory.Exists(sdkFilesPath))
            //{
            //    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
            //}
            //PrintEngineFactory.PrintEngine.Initialize();
            //label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
            //label.PrintSettings.PrinterName = PrintMachine;

            //label.Variables["CustItemNo"].SetValue(txtCustItemNo.Text.ToString());
            //label.Variables["UserNo"].SetValue(txtUserNo.Text.ToString());
            //label.Variables["Barcode_No"].SetValue(txtBarcodeNo.Text.ToString());
            //label.Variables["Barcode_Qty"].SetValue(txtBarcodeQty.Text.ToString());
            //label.Variables["OUTPUT_DATE"].SetValue(txtFinishDate.Text.ToString());
            //label.Variables["CAVITY_NO"].SetValue(txtCavityValue.Text.ToString());
            //label.Variables["TIME_NO"].SetValue(txtFinishTime.Text.ToString());
            //label.Print(1);
            //#endregion

            #region
            string labelPath = "", printMachine = "";
            using (StreamReader labelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\WhiteLensPartsticket\\LabelName.txt", Encoding.Default))
            {
                labelPath = labelPathTxt.ReadToEnd();
            }

            using (StreamReader printMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\WhiteLensPartsticket\\PrinterName.txt", Encoding.Default))
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

            label.Variables["CustItemNo"].SetValue(txtCustItemNo.Text);
            label.Variables["UserNo"].SetValue(txtUserNo.Text);
            label.Variables["Barcode_No"].SetValue(txtBarcodeNo.Text);
            label.Variables["Barcode_Qty"].SetValue(txtBarcodeQty.Text);
            label.Variables["OUTPUT_DATE"].SetValue(txtFinishDate.Text);
            label.Variables["CAVITY_NO"].SetValue(txtCavityValue.Text);
            label.Variables["TIME_NO"].SetValue(txtFinishTime.Text);
            label.Print(1);
            #endregion


        }

        private void txtMoId_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void btnLotPrint_Click(object sender, EventArgs e)
        {
            try {
                #region //成形部品票資訊
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        MoId = txtMoId.Text.ToString(),
                        BARCODE_NO = ""
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryLensPartsticketDate", jStr);
                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = TOTAL_COUNT,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });
                #endregion

                for (int i = 0; i <= TOTAL_COUNT; i++)
                {
                    #region//Get Data
                    DateTime dt = Convert.ToDateTime(jRes["qrydata"][i]["FinishDate"].ToString());
                    var TIME_NO = "";
                    if (dt.Hour >= 7 && dt.Hour < 11)
                    {
                        TIME_NO = "A(7~11點)";
                    }
                    else if (dt.Hour >= 11 && dt.Hour < 15)
                    {
                        TIME_NO = "B(11~15點)";
                    }
                    else if (dt.Hour >= 15 && dt.Hour < 19)
                    {
                        TIME_NO = "C(15~19點)";
                    }
                    else if (dt.Hour >= 19 && dt.Hour < 23)
                    {
                        TIME_NO = "E(19~23點)";
                    }
                    else if (dt.Hour >= 23 || dt.Hour < 3)
                    {
                        TIME_NO = "F(23~3點)";
                    }
                    else if (dt.Hour >= 3 && dt.Hour < 7)
                    {
                        TIME_NO = "G(3~7點)";
                    }
                    #endregion

                    #region //標籤樣式設定+印表機選擇                            
                    string LabelPath = "", LabelPathLine;
                    string PrintMachine = "", PrintMachineLine;

                    StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\WhiteLensPartsticket\\LabelName.txt", Encoding.Default);
                    while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
                    {
                        LabelPath = LabelPathLine.ToString();
                    }

                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\WhiteLensPartsticket\\PrinterName.txt", Encoding.Default);
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

                    string[] FinishDateArray = jRes["qrydata"][i]["FinishDate"].ToString().Split(' ');
                    txtFinishDate.Text = FinishDateArray[0].ToString();//日期

                    #region //標籤列印
                    label.Variables["CustItemNo"].SetValue(jRes["qrydata"][i]["MtlItemNo"].ToString());
                    label.Variables["UserNo"].SetValue(jRes["qrydata"][i]["UserNo"].ToString());
                    label.Variables["Barcode_No"].SetValue(jRes["qrydata"][i]["BarcodeNo"].ToString());
                    label.Variables["Barcode_Qty"].SetValue(jRes["qrydata"][i]["BarcodeQty"].ToString());
                    label.Variables["OUTPUT_DATE"].SetValue(FinishDateArray[0].ToString());
                    label.Variables["CAVITY_NO"].SetValue(jRes["qrydata"][i]["ItemValue"].ToString());
                    label.Variables["TIME_NO"].SetValue(TIME_NO);
                    label.Print(1);
                    #endregion

                }
            }
            catch (Exception ex)
            {
                MESConfig.LogData("白物成型部品票標籤:" + ex.ToString());
            }
        }
    }
}
