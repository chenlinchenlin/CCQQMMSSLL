using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;

namespace DX_QMS.SystemConfig
{
    public partial class QMSModel : DevExpress.XtraEditors.XtraForm
    {
        public QMSModel()
        {
            InitializeComponent();
        }

        private void QMSModel_Load(object sender, EventArgs e)
        {

           // QMS_Permission();
          //  gridView.ExpandAllGroups();
        }

        void QMS_Permission()
        {
            string sql = @"select module 模块,moduleGroup 组别,menuname 菜单,from QMS_AutoExtractMenu ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            treeList.DataSource = dt;

            treeList.KeyFieldName = "菜单";
            treeList.ParentFieldName = "模块"; 
            treeList.DataSource = dt;
            treeList.ExpandAll(); 

        }


        private void btnSave_Click(object sender, EventArgs e)
        {

        }





    }
}