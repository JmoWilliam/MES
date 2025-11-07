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
    public partial class SunnyLabel : Form
    {
        private ILabel label;
        public SunnyLabel()
        {
            InitializeComponent();
            txtSupplierNo.Text = "A10"; //供應商 A10
        }

        private void txtMoId_KeyDown(object sender, KeyEventArgs e)
        {
            try {
                if (e.KeyCode == Keys.Enter) {
                    
                    #region //選擇 部番+圖號
                    JObject jObj = JObject.FromObject(new
                    {
                        PARAMETER = JObject.FromObject(new
                        {
                            MoId = txtMoId.Text.ToString()
                        })
                    });
                    string jStr = JsonConvert.SerializeObject(jObj);
                    DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryWipRdInfo", jStr);
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
                    string MTL_NAME = jRes["qrydata"][0]["MtlItemName"].ToString();
                    string CUSTOMER_DWG_NO = jRes["qrydata"][0]["CustomerDwgNo"].ToString();
                    string editionString = jRes["qrydata"][0]["Edition"].ToString();
                    int CUSTOMER_EDITION = !string.IsNullOrEmpty(editionString) ? int.Parse(editionString) : 0;
                    string CUSTOMER_NO = "";
                    if (CUSTOMER_DWG_NO.Contains("M"))
                    {
                        string str = CUSTOMER_DWG_NO.TrimStart('M');
                        string[] CUSTOMER_NO_ARRAY = str.Split('-');
                        CUSTOMER_NO = CUSTOMER_NO_ARRAY[0].ToString();
                    }
                    else
                    {
                        CUSTOMER_NO = CUSTOMER_DWG_NO;
                    }

                    string[] ProductInfo = MTL_NAME.Split(' ');
                    if (ProductInfo.Length == 3)
                    {
                        string ProductType = ProductInfo[1];
                        string ProductName = ProductInfo[2];
                        switch (ProductType)
                        {
                            case "公模仁":
                                txtProductType.Text = "A.可動側模芯";
                                txtProductType.ReadOnly = false;
                                string[] Name = ProductName.Split('-');
                                txtSunnyName.Text = Name[0] + "-" + Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            case "母模仁":
                                txtProductType.Text = "B.固定側模芯";
                                txtProductType.ReadOnly = false;
                                string[] B_Name = ProductName.Split('-');
                                txtSunnyName.Text = B_Name[0] + "-" + B_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            case "公套筒":
                            case "公襯套":
                                txtProductType.Text = "D.可動側襯套";
                                txtProductType.ReadOnly = false;
                                string[] D_Name = ProductName.Split('-');
                                txtSunnyName.Text = D_Name[0] + "-" + D_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            case "母套筒":
                            case "母襯套":
                                txtProductType.Text = "C.固定側襯套";
                                txtProductType.ReadOnly = false;
                                string[] C_Name = ProductName.Split('-');
                                txtSunnyName.Text = C_Name[0] + "-" + C_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            default:
                                txtProductType.ReadOnly = false;
                                string[] E_Name = ProductName.Split('-');
                                txtSunnyName.Text = E_Name[0] + "-" + E_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                        }
                        Check();
                    }
                    else {
                        string ProductType = ProductInfo[0];
                        string ProductName = ProductInfo[1];
                        switch (ProductType)
                        {
                            case "公模仁":
                                txtProductType.Text = "A.可動側模芯";
                                txtProductType.ReadOnly = false;
                                string[] Name = ProductName.Split('-');
                                txtSunnyName.Text = Name[0] + "-" + Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            case "母模仁":
                                txtProductType.Text = "B.固定側模芯";
                                txtProductType.ReadOnly = false;
                                string[] B_Name = ProductName.Split('-');
                                txtSunnyName.Text = B_Name[0] + "-" + B_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            case "公套筒":
                            case "公襯套":
                                txtProductType.Text = "D.可動側襯套";
                                txtProductType.ReadOnly = false;
                                string[] D_Name = ProductName.Split('-');
                                txtSunnyName.Text = D_Name[0] + "-" + D_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            case "母套筒":
                            case "母襯套":
                                txtProductType.Text = "C.固定側襯套";
                                txtProductType.ReadOnly = false;
                                string[] C_Name = ProductName.Split('-');
                                txtSunnyName.Text = C_Name[0] + "-" + C_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                            default:
                                txtProductType.ReadOnly = false;
                                string[] E_Name = ProductName.Split('-');
                                txtSunnyName.Text = E_Name[0] + "-" + E_Name[1];
                                if (CUSTOMER_DWG_NO.Contains("X01"))
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "X01";
                                    }
                                }
                                else
                                {
                                    if (CUSTOMER_DWG_NO.Contains("CP"))
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION + "CP";
                                    }
                                    else
                                    {
                                        txtDrawingsNo.Text = CUSTOMER_NO + CUSTOMER_EDITION;
                                    }
                                }
                                break;
                                
                        }
                        Check();

                    }
                }
            }
            catch (Exception ex)
            {
                MESConfig.LogData("舜宇包裝標籤:" + ex.ToString());
            }
        }

        private void Check()
        {
            txtSunnyName.ReadOnly = false;
            txtModelNo.ReadOnly = false;
            txtModelNo.Focus();
        }

        private void txtModelNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtNum.ReadOnly = false;
                txtNum.Focus();
            }
        }

        private void txtNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtProject.ReadOnly = false;
                txtProject.Focus();
            }
        }

        private void txtProject_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtDrawingsNo.ReadOnly = false;
                txtDrawingsNo.Focus();
            }
        }

        private void txtDrawingsNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtShipDate.Focus();
                txtShipDate.ReadOnly = false;
                txtRemarks.ReadOnly = false;
            }
        }

        private void txtShipDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtRemarks.Focus();
            }
        }

        private void PrintButton_Click(object sender, EventArgs e)
        {
            try {
            string ModelNo, SunnyName, Project, Num, ProductType, Supplier, DrawingsNo, barcode;
            #region //一維條碼 資訊串接
            //模號+部番+方案+數量+類別+供應商+圖號
            if (txtModelNo.Text.ToString() != "")
            {
                ModelNo = txtModelNo.Text.ToString();
            }
            else { ModelNo = "/"; }

            if (txtSunnyName.Text.ToString() != "")
            {
                if (txtModelNo.Text.ToString() != "")
                {
                    SunnyName = "/" + txtSunnyName.Text.ToString();
                }
                else
                {
                    SunnyName = txtSunnyName.Text.ToString();
                }
            }
            else { SunnyName = "/"; }
            if (txtProject.Text.ToString() != "")
            {
                Project = "/" + txtProject.Text.ToString();
            }
            else { Project = "/"; }
            if (txtNum.Text.ToString() != "")
            {
                Num = "/" + txtNum.Text.ToString();
            }
            else { Num = "/"; }
            if (txtProductType.Text.ToString() != "")
            {
                if (txtProductType.Text.ToString() == "A.可動側模芯" || txtProductType.Text.ToString() == "A")
                {
                    ProductType = "/A";
                }
                else if (txtProductType.Text.ToString() == "B.固定側模芯" || txtProductType.Text.ToString() == "B")
                {
                    ProductType = "/B";
                }
                else if (txtProductType.Text.ToString() == "C.固定側襯套" || txtProductType.Text.ToString() == "C")
                {
                    ProductType = "/C";
                }
                else if (txtProductType.Text.ToString() == "D.可動側襯套" || txtProductType.Text.ToString() == "D")
                {
                    ProductType = "/D";
                }
                else
                {
                    ProductType = "/";
                }
            }
            else { ProductType = "/"; }
            if (txtSupplierNo.Text.ToString() != "")
            {
                Supplier = "/" + txtSupplierNo.Text.ToString();
            }
            else { Supplier = "/"; }
            if (txtDrawingsNo.Text.ToString() != "")
            {
                DrawingsNo = "/" + txtDrawingsNo.Text.ToString();
            }
            else { DrawingsNo = "/"; }
            barcode = ModelNo + SunnyName + Project + Num + ProductType + Supplier + DrawingsNo;
            #endregion

            #region //標籤樣式設定+印表機選擇                            
            string LabelPath = "", LabelPathLine;
            string PrintMachine = "", PrintMachineLine;

            StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyLabel\\LabelName.txt", Encoding.Default);
            while ((LabelPathLine = LabelPathTxt.ReadLine()) != null)
            {
                LabelPath = LabelPathLine.ToString();
            }

            StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\SunnyLabel\\PrinterName.txt", Encoding.Default);
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

            #region //標籤列印
            label.Variables["BarCode"].SetValue(barcode);
            label.Variables["date"].SetValue(txtShipDate.Text.ToString());
            label.Variables["note"].SetValue(txtRemarks.Text.ToString());
            label.Print(1);
            #endregion

            }
            catch (Exception ex)
            {
                MESConfig.LogData("舜宇外箱標籤列印異常:" + ex.ToString());
            }
        }
    }
}
