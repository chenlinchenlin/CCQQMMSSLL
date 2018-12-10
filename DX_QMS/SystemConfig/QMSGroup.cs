using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DX_QMS.Common;

namespace DX_QMS.SystemConfig
{
    public partial class QMSGroup : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public QMSGroup()
        {
            InitializeComponent();
        }

        private void QMSGroup_Load(object sender, EventArgs e)
        {

        }

        private void sBtngroupnameadd_Click(object sender, EventArgs e)
        {
            string groupname = txtgroupname.Text;
            string groupdescribe = txtgroupdescribe.Text;
            if (groupname == "")
            {
                MessageBox.Show("组别不能为空","提醒",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return;
            }
            string testsql = @" select * from QMS_Group where groupname='"+groupname+ "' order by groupid ";
            DataTable dt = DbAccess.SelectBySql(testsql).Tables[0];
            if (dt != null || dt.Rows.Count > 0)
            {
                MessageBox.Show("该组别已存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string sql = @" insert into QMS_Group( groupname ,groupdescribe ) values( '"+ groupname + " ','"+groupdescribe+ "' ) ";
            bool flat = DbAccess.ExecuteSql(sql);
            if(flat)
                MessageBox.Show("添加成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
             else
                MessageBox.Show("添加失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            sBtngroupnameselect_Click(sender,e);

        }

        private void sBtngroupnameselect_Click(object sender, EventArgs e)
        {
            string groupname, groupdescribe;
            groupname = txtgroupname.Text;
            groupdescribe = txtgroupdescribe.Text;
            string where = "where 1=1";
            if (!string.IsNullOrEmpty(groupname))
            {
                where += " and groupname ='" + groupname + "' ";
            }
            if (!string.IsNullOrEmpty(groupdescribe))
            {
                where += " and groupdescribe like '%" + groupdescribe + "%'";
            }
            string sql = @" select * from QMS_Group ";
            sql += where + " order by groupid ";
            Controlgroupname.DataSource = DbAccess.SelectBySql(sql).Tables[0];
        }

        private void sBtngroupnamedelete_Click(object sender, EventArgs e)
        {
            int i = gridgroupname.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选择需要删除的组别", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string groupname = txtgroupname.Text;
            string groupdescribe = txtgroupdescribe.Text;
            if (groupname == "")
            {
                MessageBox.Show("组别不能为空", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string sql = @" delete QMS_Group where groupname='"+groupname+ "' and groupdescribe='"+ groupdescribe + "'";
            bool flat = DbAccess.ExecuteSql(sql);
            if(flat)
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            sBtngroupnameselect_Click(sender, e);
        }

        private void sBtngroupnameupdate_Click(object sender, EventArgs e)
        {
            txtgroupname.Text = "";
            txtgroupdescribe.Text = "";
        }

        /*

        private void sBtnmenuselect_Click(object sender, EventArgs e)
        {
            string menuname;
            menuname = txtmenu.Text;
            string where = "where 1=1";
            if (!string.IsNullOrEmpty(menuname))
            {
                where += " and menuname like '%" + menuname + "%' ";
            }
            string sql = @" select * from QMS_Menu ";
            sql += where + " order by updatetime ";
            ControlMenu.DataSource = DbAccess.SelectBySql(sql).Tables[0];

        }

        private void gridMenu_Click(object sender, EventArgs e)
        {
            int i = gridMenu.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DataTable dt = ControlMenu.DataSource as DataTable;
            if (dt == null && dt.Rows.Count < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string menuname = gridMenu.GetFocusedRowCellValue("menuname").ToString();
            txtmenu.Text = menuname; 
        }
        */

        private void gridgroupname_Click(object sender, EventArgs e)
        {
            int i = gridgroupname.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DataTable dt = Controlgroupname.DataSource as DataTable;
            if (dt == null && dt.Rows.Count < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string groupname = gridgroupname.GetFocusedRowCellValue("groupname").ToString();
            string groupdescribe = gridgroupname.GetFocusedRowCellValue("groupdescribe").ToString();
            txtgroupname.Text = groupname;
            txtgroupdescribe.Text = groupdescribe;
        }

        /*
        private void sBtnPermissionadd_Click(object sender, EventArgs e)
        {
            string groupname = txtgroupname.Text;
            string menu = txtmenu.Text;
            string check = "否";
            if (groupname == "")
            {
                MessageBox.Show("组别名称不能为空", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (menu == "")
            {
                MessageBox.Show("菜单名称不能为空", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (checkPermission.Checked)
                check = "是";

            string testsql = @" select * from QMS_PermissionMenu where menuname ='"+ menu + "' and groupname='"+ groupname + "' order by Rightid";
            DataTable dt = DbAccess.SelectBySql(testsql).Tables[0];
            if (dt != null || dt.Rows.Count > 0)
            {
                MessageBox.Show("该权限已存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string sql = @" insert into QMS_PermissionMenu( menuname,groupname,permissions) values( '" + menu + " ','" + groupname + "','"+check+ "')";
            bool flat = DbAccess.ExecuteSql(sql);
            if (flat)
                MessageBox.Show("新增成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("新增失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            sBtnPermissionselect_Click(sender,e);
        }

        private void sBtnPermissionselect_Click(object sender, EventArgs e)
        {
            string groupname, menu;
            string check = "否";
            groupname = txtgroupname.Text;
            menu = txtmenu.Text;
            if (checkPermission.Checked)
                check = "是";
            string where = "where 1=1";
            if (!string.IsNullOrEmpty(groupname))
            {
                where += " and groupname ='" + groupname + "' ";
            }
            if (!string.IsNullOrEmpty(menu))
            {
                where += " and menuname ='" + menu + "'";
            }
            where += " and permissions ='"+ check + "'";
            string sql = @" select * from QMS_PermissionMenu ";
            sql += where + " order by Rightid ";
            ControlPermission.DataSource = DbAccess.SelectBySql(sql).Tables[0];

        }
        private void sBtnPermissiondelete_Click(object sender, EventArgs e)
        {
            int i = gridPermission.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选择需要删除的权限", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string menuname = gridPermission.GetFocusedRowCellValue("menuname").ToString();
            string groupname = gridPermission.GetFocusedRowCellValue("groupname").ToString();
            string permissions = gridPermission.GetFocusedRowCellValue("permissions").ToString();
            string sql = @" delete QMS_PermissionMenu where groupname='" +groupname+ "' and menuname='" +menuname+ "' and permissions='"+permissions+ "'";
            bool flat = DbAccess.ExecuteSql(sql);
            if (flat)
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            sBtnPermissionselect_Click(sender, e);
        }

        private void gridPermission_Click(object sender, EventArgs e)
        {
            int i = gridPermission.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DataTable dt = ControlPermission.DataSource as DataTable;
            if (dt == null && dt.Rows.Count < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string menuname = gridPermission.GetFocusedRowCellValue("menuname").ToString();
            string groupname = gridPermission.GetFocusedRowCellValue("groupname").ToString();
            string check = gridPermission.GetFocusedRowCellValue("permissions").ToString();
            txtmenu.Text = menuname;
            txtgroupname.Text = groupname;
            if (check == "是")
                checkPermission.Checked = true;
            else
                checkPermission.Checked = false;
        }

        private void sBtnPermissionupdate_Click(object sender, EventArgs e)
        {
            int i = gridPermission.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选择需要修改的权限", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string menuname = gridPermission.GetFocusedRowCellValue("menuname").ToString();
            string groupname = gridPermission.GetFocusedRowCellValue("groupname").ToString();
            string check = "否";
            if (checkPermission.Checked)
                check = "是";
            string sql = @" update QMS_PermissionMenu set menuname ='"+ menuname + "' and groupname ='"+ groupname + "' and permissions='"+ check + "'";
            bool flat = DbAccess.ExecuteSql(sql);
            if (flat)
                MessageBox.Show("修改成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("修改失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            sBtnPermissionselect_Click(sender, e);

        }


        */
    }
}