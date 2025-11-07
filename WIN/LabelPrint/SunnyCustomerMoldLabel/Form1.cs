using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using System.Collections.Generic;
using NiceLabel.SDK;
using MESIII;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Timers;
using Timer = System.Windows.Forms.Timer;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json;
using System.Text;
//-----------------------------------------------------------------------
// <copyright file="Form1.cs" company="Euro Plus">
//     Copyright © Euro Plus 2014.
// </copyright>
// <summary>This is the Form1 class.</summary>
//-----------------------------------------------------------------------
namespace SDKSimpleApp{

    public partial class Form1 : Form
    {
        public static Timer timer = new Timer();
        public static ISDClient client = new ISDClient(0);        
        private ILabel label, labe2;
        string PRINT_PERMISSION = "-1"; //預設列印權限 Y:可以全部印;N:只能印部分
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
        public Form1()
        {
            this.InitializeComponent();
            TxUser.ReadOnly = false;            

            StreamReader sr = new StreamReader("D:\\SunnyCustomerMoldLabel\\Template\\PrinterName.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                PrinterName = line.ToString();
            }

            StreamReader txLabelNamesr = new StreamReader("D:\\SunnyCustomerMoldLabel\\Template\\LabelName.txt", Encoding.Default);
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
        private void TxBarCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (TxUser.Text.ToString()!="") {
                    if (TxBarCode.Text.ToString() == "")
                    {
                        MessageBox.Show("★模仁條碼不能為空!!");
                    }
                    else
                    {
                        if (TxBarCode.Text.ToString() != "")
                        {
                            string BARCODE = TxBarCode.Text.ToString();
                            if (PRINT_PERMISSION == "Y")//判斷標籤列印權限
                            {
                                #region //Request
                                JObject jObj = JObject.FromObject(new
                                {
                                    WIP_TYPE = JObject.FromObject(new
                                    {
                                        BARCODE,
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
                                DataSet oDS = client.ctEnumerateData("MFGSO.QryLabel004", jStr);
                                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                                int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                                #endregion

                                #region //Response
                                JObject jRes = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "OK",
                                    table_rows = nTotalCount,
                                    qrydata = JsonConvert.DeserializeObject(sData)
                                });
                                #endregion 

                                string ITEM_BARCODE_INFO = jRes["qrydata"][0]["ITEM_BARCODE"].ToString();
                                string[] WIP_NAME = ITEM_BARCODE_INFO.Split('-');
                                string check = WIP_NAME[0].ToString();
                                Boolean isN = IsNumeric(WIP_NAME[0].ToString());
                                string ITEM_BARCODE_RESULT = "";
                                if (isN == true)
                                {
                                    ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO.Replace(check + "-", " ");
                                }
                                else
                                {
                                    ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO;
                                }
                                if (nTotalCount == 0)
                                {
                                    MessageBox.Show("☆模仁條碼不存在!!");
                                    TxBarCode.Text = "";
                                }
                                else
                                {
                                    BARCODE_NO = jRes["qrydata"][0]["BARCODE_NO"].ToString();
                                    string ITEM_BARCODE = "";
                                    int ITEM_BARCODE_LAST = int.Parse(ITEM_BARCODE_RESULT.Split('-').Last()); //取最後的數值
                                    if (ITEM_BARCODE_LAST > 10001)//去除重複刻碼
                                    {
                                        string[] str = ITEM_BARCODE_RESULT.Split('-');
                                        str = str.Take(str.Length - 1).ToArray();
                                        ITEM_BARCODE = String.Join("-", str).ToString();
                                    }
                                    else
                                    {
                                        ITEM_BARCODE = ITEM_BARCODE_RESULT;
                                    }
                                    #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                                    labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                                    labe2.PrintSettings.PrinterName = PrinterName;

                                    labe2.Variables["WIP_NAME"].SetValue(ITEM_BARCODE);
                                    labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                                    labe2.Print(1);
                                    #endregion
                                }
                                TxBarCode.Text = "";//清空 
                                TxUser.Text = "";
                                TxUser.ReadOnly = false;
                                TxUser.Focus();
                            }
                            else
                            {
                                //確認QA_LABEL_LOG是否有印過
                                #region //Request
                                JObject jObj = JObject.FromObject(new
                                {
                                    QA_LABEL_LOG = JObject.FromObject(new
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
                                DataSet oDS = client.ctEnumerateData("MFGSO.QryLabel011", jStr);
                                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                                int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                                #endregion

                                #region //Response
                                JObject jRes = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "OK",
                                    table_rows = nTotalCount,
                                    qrydata = JsonConvert.DeserializeObject(sData)
                                });
                                #endregion
                                
                                if (nTotalCount > 0)
                                {
                                    //有值;跳警訊
                                    MessageBox.Show("此模仁標籤已印過");
                                }
                                else
                                {
                                    //無值;印標籤
                                    #region //Request
                                    jObj = JObject.FromObject(new
                                    {
                                        WIP_TYPE = JObject.FromObject(new
                                        {
                                            BARCODE,
                                            PRINT_PERMISSION
                                        }),
                                        PAGE_INFO = JObject.FromObject(new
                                        {
                                            PAGE_INDEX,
                                            PAGE_SIZE,
                                            ORDER_BY
                                        })
                                    });
                                    jStr = JsonConvert.SerializeObject(jObj);
                                    oDS = client.ctEnumerateData("MFGSO.QryLabel004", jStr);
                                    sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                                    nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
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

                                    string ITEM_BARCODE_INFO = jRes["qrydata"][0]["ITEM_BARCODE"].ToString();
                                    int BARCODE_ID = Convert.ToInt32(jRes["qrydata"][0]["BARCODE_ID"].ToString());
                                    int MFG_ID = Convert.ToInt32(jRes["qrydata"][0]["MFG_ID"].ToString());
                                    string[] WIP_NAME = ITEM_BARCODE_INFO.Split('-');
                                    string check = WIP_NAME[0].ToString();
                                    Boolean isN = IsNumeric(WIP_NAME[0].ToString());
                                    string ITEM_BARCODE_RESULT = "";
                                    if (isN == true)
                                    {
                                        ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO.Replace(check + "-", " ");
                                    }
                                    else
                                    {
                                        ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO;
                                    }

                                    if (nTotalCount == 0)
                                    {
                                        MessageBox.Show("☆模仁條碼不存在!!");
                                        TxBarCode.Text = "";
                                    }
                                    else
                                    {
                                        BARCODE_NO = jRes["qrydata"][0]["BARCODE_NO"].ToString();
                                        string ITEM_BARCODE = "";
                                        int ITEM_BARCODE_LAST = int.Parse(ITEM_BARCODE_RESULT.Split('-').Last()); //取最後的數值
                                        if (ITEM_BARCODE_LAST > 10001)//去除重複刻碼
                                        {
                                            string[] str = ITEM_BARCODE_RESULT.Split('-');
                                            str = str.Take(str.Length - 1).ToArray();
                                            ITEM_BARCODE = String.Join("-", str).ToString();
                                        }
                                        else
                                        {
                                            ITEM_BARCODE = ITEM_BARCODE_RESULT;
                                        }
                                        #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                                        labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                                        labe2.PrintSettings.PrinterName = PrinterName;
                                        labe2.Variables["WIP_NAME"].SetValue(ITEM_BARCODE);
                                        labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                                        labe2.Print(1);
                                        #endregion
                                        TxBarCode.Text = "";//清空 
                                        TxUser.Text = "";
                                        TxUser.ReadOnly = false;
                                    }
                                    if (PRINT_PERMISSION == "N")
                                    {
                                        //新增QA_LABEL_LOG紀錄 
                                        string TxType = "A";
                                        int LABEL_TYPE = 2; //1:新鉅科技出貨標籤-2:舜宇出貨標籤-3:其他出貨標籤
                                        //新寫法
                                        #region //Request
                                        jObj = JObject.FromObject(new
                                        {
                                            QA_LABEL_LOG_INPUT = JObject.FromObject(new
                                            {
                                                MFG_ID,
                                                BARCODE_ID,
                                                LABEL_TYPE,
                                                USER_ID,
                                            })
                                        });
                                        jStr = JsonConvert.SerializeObject(jObj);

                                        int nResult = -1;
                                        switch (TxType)
                                        {
                                            case "A": nResult = client.ctPostTxact("MFGSO.TxQA_LabelLog", jStr, TxTypeConsts.TxTypeAddNew); break;
                                            case "D": nResult = client.ctPostTxact("MFGSO.TxQA_LabelLog", jStr, TxTypeConsts.TxTypeDelete); break;
                                            case "U": nResult = client.ctPostTxact("MFGSO.TxQA_LabelLog", jStr, TxTypeConsts.TxTypeUpdate); break;
                                            default: throw new Exception("transaction type error!");
                                        }
                                        #endregion 
                                    }
                                }
                                TxBarCode.Text = "";//清空 
                                TxUser.Text = "";
                                TxUser.ReadOnly = false;
                                TxUser.Focus();

                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("請輸入員工工號");
                    TxBarCode.Text = "";//清空 
                    TxUser.Text = "";
                    txtUserPassword.Text = "";//清空
                    TxUser.ReadOnly = false;
                    TxUser.Focus();
                }
            }
        }
        private void txMFG_KeyDown(object sender, KeyEventArgs e)
        {         
            if (e.KeyCode == Keys.Enter)
            {
                if (TxUser.Text.ToString() != "")
                {
                    if (PRINT_PERMISSION == "Y")//判斷標籤列印權限
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
                                int MFG = int.Parse(txMFG.Text.ToString());
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
                                    oDS = client.ctEnumerateData("MFGSO.QryLabel007", jStr);
                                }
                                else
                                {
                                    oDS = client.ctEnumerateData("MFGSO.QryLabel012", jStr);
                                }
                                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                                int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                                #endregion

                                #region //Response
                                JObject jRes = JObject.FromObject(new
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
                                    TxBarCode.Text = "";
                                }
                                else
                                {
                                    string ITEM_BARCODE;
                                    for (int i = 0; i < nTotalCount; i++)
                                    {
                                        timer.Interval = 2000;
                                        timer.Start();
                                        BARCODE_NO = jRes["qrydata"][i]["BARCODE_NO"].ToString();
                                        string ITEM_BARCODE_INFO = jRes["qrydata"][i]["ITEM_BARCODE"].ToString();
                                        string[] WIP_NAME = ITEM_BARCODE_INFO.Split('-');
                                        string check = WIP_NAME[0].ToString();
                                        Boolean isN = IsNumeric(WIP_NAME[0].ToString());
                                        string ITEM_BARCODE_RESULT = "";
                                        if (isN == true)
                                        {
                                            ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO.Replace(check + "-", " ");
                                        }
                                        else
                                        {
                                            ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO;
                                        }

                                        int ITEM_BARCODE_LAST = int.Parse(ITEM_BARCODE_RESULT.Split('-').Last()); //取最後的數值
                                        if (ITEM_BARCODE_LAST > 10001)//去除重複刻碼
                                        {
                                            string[] str = ITEM_BARCODE_RESULT.Split('-');
                                            str = str.Take(str.Length - 1).ToArray();
                                            ITEM_BARCODE = String.Join("-", str).ToString();
                                        }
                                        else
                                        {
                                            ITEM_BARCODE = ITEM_BARCODE_RESULT;
                                        }
                                        #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                                        labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                                        labe2.PrintSettings.PrinterName = PrinterName;
                                        labe2.Variables["WIP_NAME"].SetValue(ITEM_BARCODE);
                                        labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                                        labe2.Print(1);
                                        #endregion

                                        timer.Stop();
                                    }

                                }
                            }
                        }
                        txMFG.Text = "";//清空
                        TxUser.Text = "";//清空
                        txtUserPassword.Text = "";//清空
                        TxUser.ReadOnly = false;
                        TxUser.Focus();
                    }
                    else
                    {
                        if (txMFG.Text.ToString() == "")
                        {
                            MessageBox.Show("★製令條碼不能為空!!");
                        }
                        else
                        {                            
                            string ITEM_BARCODE;
                            int BARCODE_ID = -1;
                            string MFG = txMFG.Text.ToString();
                            if (txMFG.Text.ToString() != "")
                            {
                                #region //Request
                                JObject jObj = JObject.FromObject(new
                                {
                                    WIP_TYPE = JObject.FromObject(new
                                    {
                                        USER_ID,
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
                                    oDS = client.ctEnumerateData("MFGSO.QryLabel007", jStr);
                                }
                                else
                                {
                                    oDS = client.ctEnumerateData("MFGSO.QryLabel012", jStr);
                                }
                                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                                int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                                #endregion

                                #region //Response
                                JObject jRes = JObject.FromObject(new
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
                                    TxBarCode.Text = "";
                                }
                                else
                                {
                                    for (int i = 0; i < nTotalCount; i++)
                                    {
                                        timer.Interval = 2000;
                                        timer.Start();
                                        BARCODE_ID = Convert.ToInt32(jRes["qrydata"][i]["BARCODE_ID"].ToString());
                                        int MFG_ID = Convert.ToInt32(jRes["qrydata"][i]["MFG_ID"].ToString());
                                        BARCODE_NO = jRes["qrydata"][i]["BARCODE_NO"].ToString();
                                        string ITEM_BARCODE_INFO = jRes["qrydata"][i]["ITEM_BARCODE"].ToString();
                                        string[] WIP_NAME = ITEM_BARCODE_INFO.Split('-');
                                        string check = WIP_NAME[0].ToString();
                                        Boolean isN = IsNumeric(WIP_NAME[0].ToString());
                                        string ITEM_BARCODE_RESULT = "";
                                        if (isN == true)
                                        {
                                            ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO.Replace(check + "-", " ");
                                        }
                                        else
                                        {
                                            ITEM_BARCODE_RESULT = ITEM_BARCODE_INFO;
                                        }

                                        int ITEM_BARCODE_LAST = int.Parse(ITEM_BARCODE_RESULT.Split('-').Last()); //取最後的數值
                                        if (ITEM_BARCODE_LAST > 10001)//去除重複刻碼
                                        {
                                            string[] str = ITEM_BARCODE_RESULT.Split('-');
                                            str = str.Take(str.Length - 1).ToArray();
                                            ITEM_BARCODE = String.Join("-", str).ToString();
                                        }
                                        else
                                        {
                                            ITEM_BARCODE = ITEM_BARCODE_RESULT;
                                        }
                                        #region //標籤樣式設定+印表機選擇 //客戶標籤                            
                                        labe2 = PrintEngineFactory.PrintEngine.OpenLabel(LabelName);
                                        labe2.PrintSettings.PrinterName = PrinterName;
                                        labe2.Variables["WIP_NAME"].SetValue(ITEM_BARCODE);
                                        labe2.Variables["BARCODE_NO"].SetValue(BARCODE_NO);
                                        labe2.Print(1);
                                        #endregion
                                        timer.Stop();
                                        if (PRINT_PERMISSION == "N")
                                        {
                                            //新增QA_LABEL_LOG紀錄 
                                            string TxType = "A";
                                            int LABEL_TYPE = 2; //1:新鉅科技出貨標籤-2:舜宇出貨標籤-3:其他出貨標籤
                                                                //新寫法
                                            #region //Request
                                            jObj = JObject.FromObject(new
                                            {
                                                QA_LABEL_LOG_INPUT = JObject.FromObject(new
                                                {
                                                    MFG_ID,
                                                    BARCODE_ID,
                                                    LABEL_TYPE,
                                                    USER_ID
                                                })
                                            });
                                            jStr = JsonConvert.SerializeObject(jObj);

                                            int nResult = -1;
                                            switch (TxType)
                                            {
                                                case "A": nResult = client.ctPostTxact("MFGSO.TxQA_LabelLog", jStr, TxTypeConsts.TxTypeAddNew); break;
                                                case "D": nResult = client.ctPostTxact("MFGSO.TxQA_LabelLog", jStr, TxTypeConsts.TxTypeDelete); break;
                                                case "U": nResult = client.ctPostTxact("MFGSO.TxQA_LabelLog", jStr, TxTypeConsts.TxTypeUpdate); break;
                                                default: throw new Exception("transaction type error!");
                                            }
                                            #endregion
                                        }

                                    }

                                }
                            }
                        }
                        txMFG.Text = "";//清空
                        TxUser.Text = "";//清空
                        txtUserPassword.Text = "";//清空
                        TxUser.ReadOnly = false;
                        TxUser.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("請輸入員工工號");
                    txMFG.Text = "";//清空
                    TxUser.Text = "";//清空                    
                    txtUserPassword.Text = "";//清空
                    TxUser.ReadOnly = false;
                    TxUser.Focus();
                }
            }
        }

        private void TxUser_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtUserPassword.Focus();
            }
          
        }
        private void txtUserPassword_Keydown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string Password = GetEncoding001(txtUserPassword.Text.ToString());
                string USER_NO = TxUser.Text.ToString();

                #region //Request
                JObject jObj = JObject.FromObject(new
                {
                    PRINT_PERMISSION = JObject.FromObject(new
                    {
                        USER_NO,
                        Password
                    }),
                    PAGE_INFO = JObject.FromObject(new
                    {
                        PAGE_INDEX,
                        PAGE_SIZE,
                        ORDER_BY
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                oDS = client.ctEnumerateData("MFGSO.QryLabel009", jStr);
                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, JsonConvertSetting);
                int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                #endregion

                #region //Response
                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = nTotalCount,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });
                #endregion
                txtUserPassword.Text = "*********";
                //確認列印工號權限  
                string TYPE_VALUE = jRes["qrydata"][0]["TYPE_VALUE"].ToString(); //判斷權限
                if (TYPE_VALUE == "")
                {
                    PRINT_PERMISSION = "N";
                    TxUser.Text = jRes["qrydata"][0]["USER_NAME_CH"].ToString(); //帶出人名
                    TxUser.ReadOnly = false;
                    USER_ID = Convert.ToInt32(jRes["qrydata"][0]["USER_ID"].ToString());
                }
                else
                {
                    PRINT_PERMISSION = "Y";
                    TxUser.Text = jRes["qrydata"][0]["USER_NAME_CH"].ToString(); //帶出人名
                    TxUser.ReadOnly = false;
                    USER_ID = Convert.ToInt32(jRes["qrydata"][0]["USER_ID"].ToString());
                }
            }
        }
        public bool IsNumeric(String strNumber)
        {
            Regex NumberPattern = new Regex("[^0-9.-]");
            return !NumberPattern.IsMatch(strNumber);
        }
        #region //GetEncoding001 //密碼轉碼
        public static string GetEncoding001(string LOGIN_PWD)
        {
            try
            {
                LOGIN_PWD = new ISDServer().MyEncode01(LOGIN_PWD);

            }
            catch (Exception e1)
            {
                return e1.Message;
            }

            return LOGIN_PWD;
        }
        #endregion
    }
}
