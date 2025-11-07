using MESIII;
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

namespace MES2LensAssemblyLabel
{
    public partial class Form1 : Form
    {
        public static ISDClient client = new ISDClient(0);
        string PrinterName = "", BigLabelPath = "", SmallLabelPath = "";
        public static Timer timer = new Timer();
        private ILabel label;

        public Form1()
        {
            StreamReader srPrinterName = new StreamReader("D:\\JMO Projects\\JMO_Prod_New\\WIN\\LabelPrint\\MES2LensAssemblyLabel\\Template\\PrinterName.txt", Encoding.Default);
            String linePrinterName;
            while ((linePrinterName = srPrinterName.ReadLine()) != null)
            {
                PrinterName = linePrinterName.ToString();
            }

            StreamReader srBigLabelPath = new StreamReader("D:\\JMO Projects\\JMO_Prod_New\\WIN\\LabelPrint\\MES2LensAssemblyLabel\\Template\\BigLabelPath.txt", Encoding.Default);
            String lineBigLabelPath;
            while ((lineBigLabelPath = srBigLabelPath.ReadLine()) != null)
            {
                BigLabelPath = lineBigLabelPath.ToString();
            }

            StreamReader srSmallLabelPath = new StreamReader("D:\\JMO Projects\\JMO_Prod_New\\WIN\\LabelPrint\\MES2LensAssemblyLabel\\Template\\SmallLabelPath.txt", Encoding.Default);
            String lineSmallLabelPath;
            while ((lineSmallLabelPath = srSmallLabelPath.ReadLine()) != null)
            {
                SmallLabelPath = lineSmallLabelPath.ToString();
            }


            InitializeComponent();

            #region //選擇標籤
            string sParam = "<root><LENS_LABEL_TYPE>";
            sParam += "<TYPE_NO>-1</TYPE_NO>";
            sParam += "</LENS_LABEL_TYPE>";
            sParam += "<PAGE_INFO>";
            sParam += "<PAGE_INDEX></PAGE_INDEX>";
            sParam += "<PAGE_SIZE></PAGE_SIZE>";
            sParam += "<ORDER_BY></ORDER_BY>";
            sParam += "</PAGE_INFO>";
            sParam += "</root>";
            DataSet oDS = client.ctEnumerateData("MFGSO.QryLensLabelType001", sParam);
            cboLabelType.DisplayMember = "TypeName";
            cboLabelType.ValueMember = "TypeNo";
            cboLabelType.DataSource = oDS.Tables[0];
            #endregion
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

        private void txMFG_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string sParam = "<root><LENS_ECODE_INFO>";
                sParam += "<MFG_ID>" + txMFG.Text.ToString() + "</MFG_ID>";
                sParam += "<BARCODE_NO></BARCODE_NO>";
                sParam += "</LENS_ECODE_INFO>";
                sParam += "<PAGE_INFO>";
                sParam += "<PAGE_INDEX></PAGE_INDEX>";
                sParam += "<PAGE_SIZE></PAGE_SIZE>";
                sParam += "<ORDER_BY></ORDER_BY>";
                sParam += "</PAGE_INFO>";
                sParam += "</root>";
                DataSet oDS = client.ctEnumerateData("MFGSO.QryMes2LotLabel001", sParam);
                int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                if (nTotalCount == 0)
                {
                    MessageBox.Show("此製令尚未發料");
                }
                else
                {
                    if (nTotalCount % 2 == 0)
                    {
                        int i = 0;
                        for (i = i; i < nTotalCount; i++)
                        {
                            string LotBarcode_no1 = oDS.Tables[0].Rows[i]["BarcodeNo"].ToString();
                            string LotBarcode_no2 = oDS.Tables[0].Rows[i + 1]["BarcodeNo"].ToString();
                            print_lot(LotBarcode_no1, LotBarcode_no2);
                            i = i + 1;
                        }
                    }
                    else
                    {
                        int j = 0;
                        for (j = j; j < nTotalCount - 1; j++)
                        {
                            string LotBarcode_no1 = oDS.Tables[0].Rows[j]["BarcodeNo"].ToString();
                            string LotBarcode_no2 = oDS.Tables[0].Rows[j + 1]["BarcodeNo"].ToString();
                            print_lot(LotBarcode_no1, LotBarcode_no2);
                            j = j + 1;
                        }
                        int end = nTotalCount - 1;
                        print_pcs(oDS.Tables[0].Rows[end]["BarcodeNo"].ToString());
                    }

                }
            }
        }

        private void txLotBarcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txLotBarcode.Text.ToString() == "")
                {
                    MessageBox.Show("批號條碼不可為空");
                }
                else
                {
                    string sParam = "<root><LENS_ECODE_INFO>";
                    sParam += "<MFG_ID></MFG_ID>";
                    sParam += "<BARCODE_NO>"+ txLotBarcode.Text.ToString() + "</BARCODE_NO>";
                    sParam += "</LENS_ECODE_INFO>";
                    sParam += "<PAGE_INFO>";
                    sParam += "<PAGE_INDEX></PAGE_INDEX>";
                    sParam += "<PAGE_SIZE></PAGE_SIZE>";
                    sParam += "<ORDER_BY></ORDER_BY>";
                    sParam += "</PAGE_INFO>";
                    sParam += "</root>";
                    DataSet oDS = client.ctEnumerateData("MFGSO.QryMes2LotLabel001", sParam);
                    int nTotalCount = Convert.ToInt32(oDS.Tables["TABLE_ROWS"].Rows[0]["TOTAL_COUNT"]);

                    if (nTotalCount == 0)
                    {
                        MessageBox.Show("此製令尚未發料");
                    }
                    else {
                        print_pcs(txLotBarcode.Text.ToString());
                    }                  
                }
            }
        }

        public void print_pcs(string LotBarcode)
        {
            #region //標籤樣式設定+印表機選擇
            if (cboLabelType.SelectedValue.ToString() == "1")
            {
                label = PrintEngineFactory.PrintEngine.OpenLabel(BigLabelPath);
            }
            else
            {
                label = PrintEngineFactory.PrintEngine.OpenLabel(SmallLabelPath);
            }
            label.PrintSettings.PrinterName = PrinterName;
            label.Variables["LotBarcode_no1"].SetValue(LotBarcode);
            label.Variables["LotBarcode_no2"].SetValue("NULL");
            label.Print(1);
            check();//清空
            #endregion
        }
        public void print_lot(string LotBarcode_no1, string LotBarcode_no2)
        {
            #region //標籤樣式設定+印表機選擇
            if (cboLabelType.SelectedValue.ToString() == "1")
            {
                label = PrintEngineFactory.PrintEngine.OpenLabel(BigLabelPath);
            }
            else
            {
                label = PrintEngineFactory.PrintEngine.OpenLabel(SmallLabelPath);
            }
            label.PrintSettings.PrinterName = PrinterName;
            label.Variables["LotBarcode_no1"].SetValue(LotBarcode_no1);
            label.Variables["LotBarcode_no2"].SetValue(LotBarcode_no2);
            label.Print(1);
            check();//清空
            #endregion
        }
        private void check()
        {
            txMFG.Text = "";
            txLotBarcode.Text = "";
            txMFG.Focus();
        }
    }
}
