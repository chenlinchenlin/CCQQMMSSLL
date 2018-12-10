using DevExpress.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using DevExpress.XtraWaitForm;
using DevExpress.BarManager;
using System.Drawing;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors;
using System.Windows.Forms;

namespace DX_QMS.Common
{
    class BackgroundTask
    {
        private static WaitDialogForm LoadingDlgForm = null;
       
        public static void BackgroundWork(Action<object> action, object obj)
          {
              using (BackgroundWorker bw = new BackgroundWorker())
              {
                 bw.RunWorkerCompleted += (s, e) => 
                {
                    LoadingDlgForm.Close();
                    LoadingDlgForm.Dispose(); 

                 }; 
                 bw.DoWork += (s, e) =>
                 {
                     try
                     {
                         Action<object> a = action;
                         a.Invoke(obj);
                     }
                     catch { }
                 };

                bw.RunWorkerAsync();
                 LoadingDlgForm = new WaitDialogForm("数据正在查找中......", "提示");
                //// LoadingDlgForm = new WaitDialogForm("数据正在查找中......", "提示", new Size(200, 50), ParentForm);
            }
        }


        public static void Backgroundselect(Action<object> action, object obj, Form ParentForm)
        {
            using (BackgroundWorker bw = new BackgroundWorker())
            {
                bw.RunWorkerCompleted += (s, e) =>
                {
                    LoadingDlgForm.Close();
                    LoadingDlgForm.Dispose();

                };
                bw.DoWork += (s, e) =>
                {
                    try
                    {
                        Action<object> a = action;
                        a.Invoke(obj);
                    }
                    catch { }
                };

                bw.RunWorkerAsync();
                //// LoadingDlgForm = new WaitDialogForm("数据正在查找中......", "提示");
                LoadingDlgForm = new WaitDialogForm("数据正在查找中......", "提示", new Size(200, 50), ParentForm);
            }
        }


        //public static void Backgroundselect(Action<object> action, object obj , XtraForm owner)
        //{
        //    using (BackgroundWorker bw = new BackgroundWorker())
        //    {
        //        DialogResult Message ;
        //        bw.RunWorkerCompleted += (s, e) =>
        //        {
                   
        //        };
        //        bw.DoWork += (s, e) =>
        //        {
        //            try
        //            {
        //                Action<object> a = action;
        //                a.Invoke(obj);
        //            }
        //            catch { }
        //        };

        //        bw.RunWorkerAsync();
        //        Message= XtraMessageBox.Show(owner,"数据正在查找中......","提示", MessageBoxButtons.OK);

        //    }
        //}





    }
}
