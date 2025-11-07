using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintLabelSystem
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnShipmentLabel_Click(object sender, EventArgs e)
        {
            ShipmentLabelForm shipmentLabelForm = new ShipmentLabelForm();
            shipmentLabelForm.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            shipmentLabelForm.Show();
        }

        private void btnSunnyLabel_Click(object sender, EventArgs e)
        {
            SunnyLabel sunnyLabel = new SunnyLabel();
            sunnyLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            sunnyLabel.Show();
        }
        void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {            
            this.Show();
        }

        private void btnWhiteLensPartsticket_Click(object sender, EventArgs e)
        {
            WhiteLensPartsticket whiteLensPartsticket = new WhiteLensPartsticket();
            whiteLensPartsticket.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            whiteLensPartsticket.Show();
        }

        private void btnAsphericGlassLensShippingLabel_Click(object sender, EventArgs e)
        {
            AsphericGlassLensShippingLabel asphericGlassLensShippingLabel = new AsphericGlassLensShippingLabel();
            asphericGlassLensShippingLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            asphericGlassLensShippingLabel.Show();
        }

        private void btnQcNotice_Click(object sender, EventArgs e)
        {
            QcNoticeLabel qcNoticeLabel = new QcNoticeLabel();
            qcNoticeLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            qcNoticeLabel.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SunnyCustMoldLabel sunnyCustMoldLabel = new SunnyCustMoldLabel();
            sunnyCustMoldLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            sunnyCustMoldLabel.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ToolLabel toolLabel = new ToolLabel();
            toolLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            toolLabel.Show();
        }

        private void btnOtherCustomerMoldLabel_Click(object sender, EventArgs e)
        {
            OtherCustomerMoldLabel otherCustomerMoldLabel = new OtherCustomerMoldLabel();
            otherCustomerMoldLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            otherCustomerMoldLabel.Show();
        }

        private void btnNewMaxCustomerMoldLabel_Click(object sender, EventArgs e)
        {
            NewMaxCustomerMoldLabel newMaxCustomerMoldLabel = new NewMaxCustomerMoldLabel();
            newMaxCustomerMoldLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            newMaxCustomerMoldLabel.Show();
        }

        private void btnInStockLabel_Click(object sender, EventArgs e)
        {
            InboundLabel inboundLabel = new InboundLabel();
            inboundLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            inboundLabel.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            NewSunnyCustMoldLabel newSunnyCustMoldLabel = new NewSunnyCustMoldLabel();
            newSunnyCustMoldLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            newSunnyCustMoldLabel.Show();
        }

        private void btnMockClientShippingLabels_Click(object sender, EventArgs e)
        {
            MockClientShippingLabels mockClientShippingLabels = new MockClientShippingLabels();
            mockClientShippingLabels.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            mockClientShippingLabels.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MockClientShippingLabels mockClientShippingLabels = new MockClientShippingLabels();
            mockClientShippingLabels.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            mockClientShippingLabels.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            BarcodeQuant barcodeQuant = new BarcodeQuant();
            barcodeQuant.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            barcodeQuant.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SunnyJCShip sunnyJCShip = new SunnyJCShip();
            sunnyJCShip.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            sunnyJCShip.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SunnyShipLabel sunnyShipLabel = new SunnyShipLabel();
            sunnyShipLabel.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            sunnyShipLabel.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            BarcodeQuant barcodeQuant = new BarcodeQuant();            
            barcodeQuant.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            barcodeQuant.Data = "Y"; //需要帶出班別
            barcodeQuant.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            MockClientShippingLabels mockClientShippingLabels = new MockClientShippingLabels();
            mockClientShippingLabels.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            this.Hide();
            mockClientShippingLabels.Data = "Y"; //需要帶出班別
            mockClientShippingLabels.Show();
        }
    }
}
