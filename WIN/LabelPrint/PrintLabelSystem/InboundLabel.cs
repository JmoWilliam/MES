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
    public partial class InboundLabel : Form
    {
        private ILabel label;
        public InboundLabel()
        {
            InitializeComponent();

            #region //選擇 庫別
            string TypeSchema = "Label.WarehouseStorage";
            JObject jObj = JObject.FromObject(new
            {
                PARAMETER = JObject.FromObject(new
                {
                    TypeSchema
                })
            });
            string jStr = JsonConvert.SerializeObject(jObj);
            DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryType", jStr);
            cboWAREHOUSE.DisplayMember = "TypeName";
            cboWAREHOUSE.ValueMember = "TypeName";
            cboWAREHOUSE.DataSource = oDS.Tables[0];
            #endregion
        }

        private void btnPRINTER_Click(object sender, EventArgs e)
        {
            #region //標籤樣式設定+印表機選擇                            
            string LabelPath = "", LabelPathLine;
            string PrintMachine = "", PrintMachineLine;

            StreamReader LabelPathTxt = new StreamReader(@"C:\WIN\PrintLabelSystem\Template\InboundLabel\LabelName.txt", Encoding.Default);
            while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
            {
                LabelPath = LabelPathLine.ToString();
            }

            StreamReader PrintMachineTxt = new StreamReader(@"C:\WIN\PrintLabelSystem\Template\InboundLabel\PrinterName.txt", Encoding.Default);
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

            if (txtMFG_ID.Text!="") {
                //用製令列印
                #region
                string MoId = txtMFG_ID.Text.ToString();
                string BarcodeNo = TxBarCode.Text.ToString();
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        MoId,
                        BarcodeNo
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryMoProcessInfo", jStr);
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
                    label.Variables["MTL_NAME"].SetValue(txtMTL_NAME.Text.ToString());
                    label.Variables["MTL_NO"].SetValue(txtMTL_NO.Text.ToString());
                    label.Variables["WAREHOUSE"].SetValue(cboWAREHOUSE.SelectedValue.ToString());
                    label.Variables["BARCODE_NO"].SetValue(jRes["qrydata"][i]["BarcodeNo"].ToString());
                    if (cboWAREHOUSE.SelectedValue.ToString() == "待報廢倉")
                    {
                        label.Variables["STATION"].SetValue(textNgReason.Text.ToString());
                    }
                    else
                    {
                        label.Variables["STATION"].SetValue(txtProcessName.Text.ToString());
                    }
                    label.Print(1);
                }
            }
            else {
                //用條碼列印
                #region
                string MoId = txtMFG_ID.Text.ToString();
                string BarcodeNo = TxBarCode.Text.ToString();
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        MoId,
                        BarcodeNo
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryMoProcessInfoBarcode", jStr);
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

                label.Variables["MTL_NAME"].SetValue(txtMTL_NAME.Text.ToString());
                label.Variables["MTL_NO"].SetValue(txtMTL_NO.Text.ToString());
                label.Variables["WAREHOUSE"].SetValue(cboWAREHOUSE.SelectedValue.ToString());
                label.Variables["BARCODE_NO"].SetValue(jRes["qrydata"][0]["BarcodeNo"].ToString());
                if (cboWAREHOUSE.SelectedValue.ToString() == "待報廢倉")
                {
                    label.Variables["STATION"].SetValue(textNgReason.Text.ToString());
                }
                else
                {
                    label.Variables["STATION"].SetValue(txtProcessName.Text.ToString());
                }
                label.Print(1);                
            }
        }

        private void txtMFG_ID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string MoId = txtMFG_ID.Text.ToString();
                string BarcodeNo = "";
                #region
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        MoId,
                        BarcodeNo
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryMoProcessInfo", jStr);
                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                int Count = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = Count,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });
                txtMTL_NO.Text = jRes["qrydata"][0]["MtlItemNo"].ToString();
                txtMTL_NAME.Text = jRes["qrydata"][0]["MtlItemName"].ToString();               

                #endregion

                #region//顯示製令製程
                JObject MoProcessJObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        MoId
                    })
                });
                string MoProcessJStr = JsonConvert.SerializeObject(MoProcessJObj);
                DataSet MoProcessODS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryMoProcess", MoProcessJStr);
                cboSTATION.DisplayMember = "ProcessName";
                cboSTATION.ValueMember = "ProcessName";
                cboSTATION.DataSource = MoProcessODS.Tables[0];
                txtProcessName.Text=cboSTATION.SelectedValue.ToString();
                #endregion
            }
        }

        private void TxBarCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string MoId = "";
                string BarcodeNo = TxBarCode.Text.ToString();
                #region
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        MoId,
                        BarcodeNo
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryMoProcessInfoBarcode", jStr);
                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                int Count = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = Count,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });
                txtMTL_NO.Text = jRes["qrydata"][0]["MtlItemNo"].ToString();
                txtMTL_NAME.Text = jRes["qrydata"][0]["MtlItemName"].ToString();
                MoId = jRes["qrydata"][0]["MoId"].ToString();
                #endregion

                #region//顯示製令製程
                JObject MoProcessJObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        MoId,
                        BarcodeNo
                    })
                });
                string MoProcessJStr = JsonConvert.SerializeObject(MoProcessJObj);
                DataSet MoProcessODS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryMoProcess", MoProcessJStr);
                cboSTATION.DisplayMember = "ProcessName";
                cboSTATION.ValueMember = "ProcessName";
                cboSTATION.DataSource = MoProcessODS.Tables[0];
                txtProcessName.Text = cboSTATION.SelectedValue.ToString();
                #endregion
            }
        }

        private void cboSTATION_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtProcessName.Text = cboSTATION.SelectedValue.ToString();
        }
    }
}
