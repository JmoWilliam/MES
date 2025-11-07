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
    public partial class QcNoticeLabel : Form
    {
        private ILabel label;
        public QcNoticeLabel()
        {
            InitializeComponent();
        }

        private void txtQcNotice_KeyDown(object sender, KeyEventArgs e)
        {
            try {
                #region //查詢需求單資訊
                JObject jObj = JObject.FromObject(new
                {
                    PARAMETER = JObject.FromObject(new
                    {
                        QcNoticeNo = txtQcNotice.Text.ToString()
                    })
                });
                string jStr = JsonConvert.SerializeObject(jObj);
                DataSet oDS = MESConfig.clientNewMes.ctEnumerateData("APISO.QryQcNoticeInfo", jStr);
                string sData = JsonConvert.SerializeObject(oDS.Tables[0], Formatting.Indented, MESConfig.JsonConvertSetting);
                int TOTAL_COUNT = MESConfig.cmn.cap_int(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);
                JObject jRes = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    table_rows = TOTAL_COUNT,
                    qrydata = JsonConvert.DeserializeObject(sData)
                });

                TypeDesc.Text = jRes["qrydata"][0]["TypeDesc"].ToString();//單據類型
                QcNoticeNo.Text = jRes["qrydata"][0]["QcNoticeNo"].ToString();//單據編號
                WoMtlItemNo.Text = jRes["qrydata"][0]["WoMtlItemNo"].ToString();//品號
                WOMtlItemName.Text = jRes["qrydata"][0]["WOMtlItemName"].ToString();//品名
                WOMtlItemSpec.Text = jRes["qrydata"][0]["WOMtlItemSpec"].ToString();//規格
                WoErp.Text = jRes["qrydata"][0]["WoErpPrefix"].ToString()+"-"+ jRes["qrydata"][0]["WoErpNo"].ToString()+"("+ jRes["qrydata"][0]["WoSeq"].ToString() + ")";//製令
                InputQty.Text = jRes["qrydata"][0]["InputQty"].ToString();//預計產量
                DepartmentName.Text = jRes["qrydata"][0]["DepartmentName"].ToString();//需求部門
                UserName.Text = jRes["qrydata"][0]["UserName"].ToString();//建立者
                CreateDate.Text = jRes["qrydata"][0]["CreateDate"].ToString();//建立日期
                #endregion
            }
            catch (Exception ex)
            {
                MESConfig.LogData("量測需求單:" + ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {
                #region
                string labelPath = "", printMachine = "";
                using (StreamReader labelPathTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\QcNoticeLabel\\LabelName.txt", Encoding.Default))
                {
                    labelPath = labelPathTxt.ReadToEnd();
                }

                using (StreamReader printMachineTxt = new StreamReader("C:\\WIN\\PrintLabelSystem\\Template\\QcNoticeLabel\\PrinterName.txt", Encoding.Default))
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

                label.Variables["TypeDesc"].SetValue(TypeDesc.Text);
                label.Variables["QcNoticeNo"].SetValue(QcNoticeNo.Text);
                label.Variables["WoMtlItemNo"].SetValue(WoMtlItemNo.Text);
                label.Variables["WOMtlItemName"].SetValue(WOMtlItemName.Text);
                label.Variables["WOMtlItemSpec"].SetValue(WOMtlItemSpec.Text);
                label.Variables["WoErp"].SetValue(WoErp.Text);
                label.Variables["InputQty"].SetValue(InputQty.Text);
                label.Variables["DepartmentName"].SetValue(DepartmentName.Text);
                label.Variables["UserName"].SetValue(UserName.Text);
                label.Variables["CreateDate"].SetValue(CreateDate.Text);
                label.Print(1);
                #endregion
            }
            catch (Exception ex)
            {
                MESConfig.LogData("量測需求單:" + ex.ToString());
            }
        }
    }
}
