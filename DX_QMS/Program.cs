using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.UserSkins;
using DevExpress.Skins;
using DevExpress.XtraSplashScreen;
//using DX_QMS.SystemConfig;
//using DX_QMS.IQCFilePosition;
namespace DX_QMS
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            BonusSkins.Register();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            BonusSkins.Register();
            SkinManager.EnableFormSkins();
            SplashScreenManager.ShowForm((Form)null, typeof(frmSplashScreen), true, true);
            Application.DoEvents();
            System.Threading.Thread.Sleep(2000);
            SplashScreenManager.CloseForm();
            System.Threading.Thread.Sleep(500);
            Application.Run(new Login());   // Login  QMSGroup  QMSMenu   TestResultCheck
        }
    }
}
