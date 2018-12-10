using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using DevExpress.XtraBars;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class OQCinformation : DevExpress.XtraBars.Ribbon.RibbonForm  //  DevExpress.XtraEditors.XtraForm  DevExpress.XtraBars.Ribbon.RibbonForm
    {
       
        public OQCinformation()
        {
            InitializeComponent();
        }
        private void setRule()
        {
            
        }

        void bindbadclass()
        {
            string sql = @"  select distinct badclass from OQC_baditem  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtbadclass.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtbadclass.Properties.Items.Add(row["badclass"]);
            }

        }
        private void OQCinformation_Load(object sender, EventArgs e)
        {
            bindbadclass();
        }

        private void txtbadclass_SelectedIndexChanged(object sender, EventArgs e)
        {
            sBtnselect_Click(sender, e);
        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {
            string badclass = "", badphenomenon = "";
            string where = " where 1=1 ";
            badclass = txtbadclass.Text.Trim();
            badphenomenon = txtbaddescribe.Text.Trim();

            if (!string.IsNullOrEmpty(badclass))
            {
                where += " and badclass = '" + badclass + "' ";
            }
            if (!string.IsNullOrEmpty(badphenomenon))
            {
                where += " and badphenomenon = '" + badphenomenon + "' ";
            }
            string sql = @"  select badclass 不良类别,badphenomenon 不良现象,defects 缺陷定义,remarks 备注,updateuser 更新人,updatetime 更新时间 from OQC_baditem  ";

            sql += where + " ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }
        }

        private void sBtninsert_Click(object sender, EventArgs e)
        {
            string  badclass = txtbadclass.Text.Trim();
            string  badphenomenon = txtbaddescribe.Text.Trim();
            string sql = @"  select 1  from OQC_baditem where  badclass ='"+ badclass + "' and badphenomenon = '" + badphenomenon + "' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                MessageBox.Show("该不良已存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
               sql = @" insert into OQC_baditem (badclass ,badphenomenon ,defects ,remarks,updateuser,updatetime )   ";
               sql = sql + " values ('"+ badclass + "','"+ badphenomenon + "','"+ txtMAMI.Text.Trim()+ "','"+ txtremark.Text+ "','"+ Login.username+ "',GETDATE()) ";
               bool falg = DbAccess.ExecuteSql(sql);
                if (falg == true)
                {
                    MessageBox.Show("添加成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    sBtnselect_Click(sender, e);
                }
                else
                {
                    MessageBox.Show("添加失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);                  
                }

            }

        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                return;
            }
            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中要删除的项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string badclass = gridView.GetFocusedRowCellValue("不良类别").ToString();
            string badphenomenon = gridView.GetFocusedRowCellValue("不良现象").ToString();

            string sql = @" delete OQC_baditem where badclass = '" + badclass + "' and badphenomenon = '" + badphenomenon + "' ";

            bool falg = DbAccess.ExecuteSql(sql);
            if (falg == true)
            {
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sBtnselect_Click(sender, e);
            }
            else
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txtbadclass.Text = "";
            txtbaddescribe.Text = "";
            txtMAMI.Text = "";
            txtremark.Text = "";
            gridControl.DataSource = null;
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                return;
            }
            txtbadclass.Text =  gridView.GetFocusedRowCellValue("不良类别").ToString();
            txtbaddescribe.Text = gridView.GetFocusedRowCellValue("不良现象").ToString();
            txtMAMI.Text = gridView.GetFocusedRowCellValue("缺陷定义").ToString();
            txtremark.Text = gridView.GetFocusedRowCellValue("备注").ToString();
        }
    }
}