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
    public partial class ShipmentLabelForm : Form
    {
        private static ILabel label;
        public ShipmentLabelForm()
        {
            InitializeComponent();

            #region //標籤類別
            JObject jObj = JObject.FromObject(new
            {
                PARAMETER = JObject.FromObject(new
                {
                    USETYPE= "JingDingWei"
                })
            });
            string jStr = JsonConvert.SerializeObject(jObj);
            DataSet oDS = MESConfig.clientOldMes.ctEnumerateData("APISO.QryJingdingweiLabelType", jStr);
            string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
            int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
            cboLabelType.DisplayMember = "LABEL_NAME";
            cboLabelType.ValueMember = "LABEL_ID";
            cboLabelType.DataSource = oDS.Tables[0];
            #endregion

            if (cboLabelType.SelectedValue.ToString()=="10006") {
                #region //廠商
                jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        LABEL_SHIPMENTID = -1
                    })
                });
                jStr = JsonConvert.SerializeObject(jObj);
                oDS = MESConfig.clientOldMes.ctEnumerateData("APISO.QryShipmentLabel", jStr);
                sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = TOTAL_COUNT,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });

                cboVendor.DisplayMember = "LABELNAME_CN";
                cboVendor.ValueMember = "LABEL_SHIPMENTID";
                cboVendor.DataSource = oDS.Tables[0];
                #endregion

                txtLabelFlowID.Text= jRes["qrydata"][0]["LABEL_FLOWID"].ToString();
                txtLabelFormat.Text= jRes["qrydata"][0]["LABEL_FORMAT"].ToString();               
                
            }
        }

        private void cboVendor_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region //廠商
            JObject jObj = JObject.FromObject(new
            {
                PARAMETER = JObject.FromObject(new
                {
                    LABEL_SHIPMENTID = cboVendor.SelectedValue.ToString()
                })
            });
            string jStr = JsonConvert.SerializeObject(jObj);
            DataSet oDS = MESConfig.clientOldMes.ctEnumerateData("APISO.QryShipmentLabel", jStr);
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
            txtLabelFlowID.Text = jRes["qrydata"][0]["LABEL_FLOWID"].ToString();
            txtLabelFormat.Text = jRes["qrydata"][0]["LABEL_FORMAT"].ToString();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            try {
                string LabelPath = "", LabelPathLine;
                string PrintMachine = "", PrintMachineLine;

                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();

                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ShipmentLabelForm\\LabelName.txt", Encoding.Default);
                while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathLine.ToString();
                }

                StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\ShipmentLabelForm\\PrinterName.txt", Encoding.Default);
                while ((PrintMachineLine = LabelPathTxt.ReadLine()) != null)
                {
                    PrintMachine = PrintMachineLine.ToString();
                }

                int LABEL_SHIPMENTID = Convert.ToInt32(cboVendor.SelectedValue.ToString()); //出貨廠商ID
                int LABEL_FLOWID = Convert.ToInt32(txtLabelFlowID.Text.ToString()); //流水號
                string LABEL_FORMAT = txtLabelFormat.Text.ToString(); //格式
                int PIECE = Convert.ToInt32(txtPiece.Text.ToString());//列印張數

                for (int i = 0; i < PIECE; i++)
                {
                    switch (LABEL_SHIPMENTID)
                    {
                        case 1:
                            LABEL_FORMAT = LABEL_FORMAT.Replace(LABEL_FORMAT.Substring(8, 3), (LABEL_FLOWID + i).ToString("000"));
                            break;
                        case 2:
                            LABEL_FORMAT = LABEL_FORMAT.Replace(LABEL_FORMAT.Substring(3, 4), (LABEL_FLOWID + i).ToString("0000"));
                            break;
                    }

                    #region //標籤樣式設定+印表機選擇
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);//選擇標籤
                    label.PrintSettings.PrinterName = PrintMachine;//選擇印表機
                    #endregion

                    label.Variables["LABEL_FORMAT"].SetValue(LABEL_FORMAT);
                    label.Print(1);
                }

                #region//更新精定位流水號標籤                
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        LABEL_SHIPMENTID,
                        LABEL_FLOWID,
                        PIECE
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                int nResult = MESConfig.clientOldMes.ctPostTxact("APISO.TxShipmentLabel", jStr, TxTypeConsts.TxTypeUpdate);                
                #endregion
            }
            catch (Exception ex)
            {
                MESConfig.LogData("精定位標籤列印:"+ex.ToString());
            }

        }
    }
}
