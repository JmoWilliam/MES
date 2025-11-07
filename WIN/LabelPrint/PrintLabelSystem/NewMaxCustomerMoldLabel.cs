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
    public partial class NewMaxCustomerMoldLabel : Form
    {
        private ILabel label;
        string MtlItemName = "", BarcodeNo = "", ItemValue = "";
        public NewMaxCustomerMoldLabel()
        {
            InitializeComponent();
            this.label_InsertType.Visible = false;
            this.text_InsertType.Visible = false;
        }

        private void TxBarCode_KeyDown(object sender, KeyEventArgs e)
        {
            try {
                if (e.KeyCode == Keys.Enter)
                {
                    if (TxBarCode.Text.ToString() == "")
                    {
                        MessageBox.Show("★模仁條碼不能為空!!");
                    }
                    else {
                        PrintLabel(TxBarCode.Text.ToString(),"-1","");
                    }
                }
            }
            catch (Exception ex)
            {
                MESConfig.LogData("新鉅出貨標籤:" + ex.ToString());
                MessageBox.Show("新鉅出貨標籤:" + ex.ToString());
            }

        }
        public void PrintLabel(string BarcodeNo, string MoId,string WIP_TYPE)
        {
            string BARCODE = TxBarCode.Text.ToString();
            #region //
            JObject jObj = JObject.FromObject(new
            {
                PARAMETER = JObject.FromObject(new
                {
                    BarcodeNo,
                    MoId
                })
            });
            string jStr = JsonConvert.SerializeObject(jObj);
            DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryOtherCustomerMold", jStr);
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

                MtlItemName = jRes["qrydata"][i]["MtlItemName"].ToString();
                BarcodeNo = jRes["qrydata"][i]["BarcodeNo"].ToString();
                ItemValue = jRes["qrydata"][i]["ItemValue"].ToString();

                if (MtlItemName.IndexOf("公模仁") > -1)
                {
                    TxTYPE.Text = "公模仁";
                    Print(MtlItemName, BarcodeNo, TxTYPE.Text.ToString(), ItemValue.Substring(ItemValue.Length - 2));
                }
                else if (MtlItemName.IndexOf("母模仁") > -1)
                {
                    TxTYPE.Text = "母模仁";
                    Print(MtlItemName, BarcodeNo, TxTYPE.Text.ToString(), ItemValue.Substring(ItemValue.Length - 2));

                }
                else if (MtlItemName.IndexOf("上模仁") > -1)
                {
                    TxTYPE.Text = "維修品";
                    Print(MtlItemName, BarcodeNo, TxTYPE.Text.ToString(), ItemValue.Substring(ItemValue.Length - 2));

                }
                else if (MtlItemName.IndexOf("下模仁") > -1)
                {
                    TxTYPE.Text = "維修品";
                    Print(MtlItemName, BarcodeNo, TxTYPE.Text.ToString(), ItemValue.Substring(ItemValue.Length - 2));
                }
                else if (MtlItemName.IndexOf("入子") > -1)
                {
                    this.label_InsertType.Visible = true;
                    this.text_InsertType.Visible = true;                   
                    text_InsertType.Focus();
                    if (WIP_TYPE == "通光孔入子" || WIP_TYPE == "鏡室入子") {
                       Print(MtlItemName, BarcodeNo, WIP_TYPE, ItemValue.Substring(ItemValue.Length - 2));
                    }
                }
            }
        }

        private void txMFG_ID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                PrintLabel("", txMFG_ID.Text.ToString(),"");
            }
        }

        private void text_InsertType_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (text_InsertType.Text.ToString() == "")
                {
                    MessageBox.Show("入子條碼不能為空!!");
                }
                else
                {
                    switch (text_InsertType.Text.ToString()) {
                        case "S1":
                        case "通光孔入子":
                            if (txMFG_ID.Text != "")
                            {
                                //MFG列印
                                PrintLabel("", txMFG_ID.Text.ToString(), "通光孔入子");
                                //MOLD_LIST(txMFG_ID.Text.ToString());
                            }
                            else
                            {
                                TxTYPE.Text = "通光孔入子";
                                Print(MtlItemName, BarcodeNo, TxTYPE.Text.ToString(), ItemValue.Substring(ItemValue.Length - 2));
                            }
                            break;
                        case "S2":
                        case "鏡室入子":
                            if (txMFG_ID.Text != "")
                            {
                                //MFG列印
                                PrintLabel("", txMFG_ID.Text.ToString(), "鏡室入子");
                            }
                            else
                            {
                                TxTYPE.Text = "鏡室入子";
                                Print(MtlItemName, BarcodeNo, TxTYPE.Text.ToString(), ItemValue.Substring(ItemValue.Length - 2));
                            }
                            break;
                    }
                }
            }
        }
        private void Print(string WIP_NAME,string BARCODE_NO,string WIP_TYPE,string SEQ) {
            #region //標籤樣式設定+印表機選擇                            
            string LabelPath = "", LabelPathLine;

            string PrintMachine = "", PrintMachineLine;

            StreamReader LabelPathTxt = new StreamReader(@"C:\WIN\PrintLabelSystem\Template\NewMaxCustomerMoldLabel\LabelName.txt", Encoding.Default);
            while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
            {
                LabelPath = LabelPathLine.ToString();
            }

            StreamReader PrintMachineTxt = new StreamReader(@"C:\WIN\PrintLabelSystem\Template\NewMaxCustomerMoldLabel\PrinterName.txt", Encoding.Default);
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

            label.Variables["WIP_NAME"].SetValue(WIP_NAME.Split(' ')[1].ToString() + '-' + SEQ);
            label.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
            label.Variables["WIP_TYPE"].SetValue(WIP_TYPE);
            label.Print(1);
        }

    }
}
