using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NikonDigimicro_MFC101AE
{
    public class ClassSerialPort
    {
        public SerialPort serialPort = new SerialPort();
        public void initial()
        {

        }
        public void initial(string port,int baudrate, StopBits stopbits,int databits)
        {
            if (serialPort.IsOpen == false)
            {
                serialPort.PortName = port;
                serialPort.BaudRate = baudrate;
                serialPort.StopBits = stopbits;
                serialPort.DataBits = databits;
                serialPort.DtrEnable = true;
                serialPort.RtsEnable = true;
                serialPort.ReadTimeout = 1000;

                serialPort.Close();
            }
        }
        public bool open()
        {
            if (serialPort.IsOpen == false)
                serialPort.Open();
  
            return serialPort.IsOpen;
        }
        public bool close()
        {
            serialPort.Close();

            return serialPort.IsOpen;
        }
        public string sendData(string str)
        {
            serialPort.Write(str);
            String input = "";
            if (str == "RX\r\n")
                input = "Reset";
            else
                input = serialPort.ReadLine();
            return input;
        }
    }
}
