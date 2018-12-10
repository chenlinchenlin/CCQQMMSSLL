using System;
using System.Threading;
using System.Windows.Forms;

namespace DX_QMS.Common
{
    class COMScanDrive
    {
        private static Thread tdReadComData;
        public static System.IO.Ports.SerialPort sPrtScan;

        public static void StartComRead()
        {
            try
            {
                if (sPrtScan == null)
                    sPrtScan = new System.IO.Ports.SerialPort();
                else if (sPrtScan.IsOpen)
                    sPrtScan.Close();
                sPrtScan.Open();
                tdReadComData = new Thread(new ThreadStart(ReadComData));
                tdReadComData.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static void ReadComData()
        {
            sPrtScan.NewLine = "\r";
            while (sPrtScan.IsOpen)
            {
                string sData = "";
                try
                {
                    sData = sPrtScan.ReadLine();
                }
                catch
                {
                    if (sPrtScan.IsOpen)
                        sPrtScan.Close();
                    return;
                }
                if (!sData.Equals(""))
                {
                    for (int i = 0; i < sData.Length; i++)
                    {
                        SendKeys.SendWait("{" + sData[i].ToString() + "}");
                    }
                    SendKeys.SendWait("{Enter}");
                }
            }
        }

        public static void EndComRead()
        {
            if (sPrtScan.IsOpen)
            {
                sPrtScan.Close();
            }
            else
            {
                sPrtScan.Dispose();
            }
            if (tdReadComData != null)
            {
                tdReadComData.Abort();
                tdReadComData = null;
            }
        }
    }
}
