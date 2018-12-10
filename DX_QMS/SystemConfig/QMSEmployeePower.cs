using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DX_QMS.Common;

namespace DX_QMS.SystemConfig
{
    public partial class QMSEmployeePower : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public QMSEmployeePower()
        {
            InitializeComponent();
            bindData();
        }
        private void bindData()
        {

            if (Login.manager == "QE管理员")
                txtpost.Properties.Items.Add("QE");
            else if (Login.manager == "SQE管理员")
                txtpost.Properties.Items.Add("SQE");
            else if (Login.manager == "IQC管理员")
            {
                txtpost.Properties.Items.Add("IQC批次生成与检验");
                txtpost.Properties.Items.Add("IQC测试员");
            }
            else if (Login.manager == "OQC管理员")
                txtpost.Properties.Items.Add("OQC");
            else if (Login.manager == "ESD管理员")
                txtpost.Properties.Items.Add("ESD");
            else if (Login.manager == "IPQC管理员")
                txtpost.Properties.Items.Add("IPQC");
            else if (Login.manager == "IPQC&OQC管理员")
                txtpost.Properties.Items.Add("IPQC&OQC");
            else if (Login.manager == "外访组管理员")
                txtpost.Properties.Items.Add("外访组");
            else if (Login.manager == "IT管理员")
            {
                txtpost.Properties.Items.Add("QE");
                txtpost.Properties.Items.Add("SQE");
                txtpost.Properties.Items.Add("IQC批次生成与检验");
                txtpost.Properties.Items.Add("IQC测试员");
                txtpost.Properties.Items.Add("OQC");
                txtpost.Properties.Items.Add("IPQC");
                txtpost.Properties.Items.Add("NPI");
            }
        }
        private void genTree(string post)
        {
            //读取已有的权限
            treeView1.Nodes[0].Nodes.Clear();
            treeView1.BeginUpdate();

            string parentNode = "select distinct model from QMS_PermissionMenu where post='" + Login.manager + "'and permissions='是'";
            DataTable dt0 = DbAccess.SelectBySql(parentNode).Tables[0];

            if (dt0 != null && dt0.Rows.Count > 0)
            {
                foreach (DataRow dr in dt0.Rows)
                {
                    TreeNode tempNode1 = new TreeNode();
                    tempNode1.Name = dr["model"].ToString();
                    tempNode1.Text = dr["model"].ToString();
                    treeView1.Nodes[0].Nodes.Add(tempNode1);
                    tempNode1.ForeColor = Color.Blue;

                    string node2 = "select distinct modelGroup from QMS_PermissionMenu where post='" + Login.manager + "'and permissions='是'and model='" + dr["model"].ToString() + "'";
                    DataTable md = DbAccess.SelectBySql(node2).Tables[0];

                    foreach (DataRow dr2 in md.Rows)
                    {
                        TreeNode tempNode2 = new TreeNode();
                        tempNode2.Name = dr2["modelGroup"].ToString();
                        tempNode2.Text = dr2["modelGroup"].ToString();

                        treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes.Add(tempNode2);
                        treeView1.Nodes[0].Nodes[dr["model"].ToString()].Expand();


                        string node3 = "select menuName from QMS_PermissionMenu where  post='" + Login.manager + "'and permissions='是'and model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"].ToString() + "'";
                        DataTable dt3 = DbAccess.SelectBySql(node3).Tables[0];


                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            TreeNode tempNode3 = new TreeNode();
                            tempNode3.Name = dr3["menuName"].ToString();
                            tempNode3.Text = dr3["menuName"].ToString();

                            treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes.Add(tempNode3);
                            treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Expand();

                            string queryPower = "select hasQuery,hasInsert,hasUpdate,hasDelete,hasPrint,hasAudit from QMS_PermissionMenu where post='" + Login.manager + "'and permissions='是'and model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"].ToString() + "'and menuName='" + tempNode3.Text + "'";
                            DataTable dt_power = DbAccess.SelectBySql(queryPower).Tables[0];

                            Dictionary<string, string> dicTemp = new Dictionary<string, string>();
                            if (bool.Parse(dt_power.Rows[0]["hasQuery"].ToString()) == true) dicTemp.Add("Select", "查看");
                            if (bool.Parse(dt_power.Rows[0]["hasInsert"].ToString()) == true) dicTemp.Add("Insert", "增加");
                            if (bool.Parse(dt_power.Rows[0]["hasUpdate"].ToString()) == true) dicTemp.Add("Update", "修改");
                            if (bool.Parse(dt_power.Rows[0]["hasDelete"].ToString()) == true) dicTemp.Add("Delete", "删除");
                            if (bool.Parse(dt_power.Rows[0]["hasPrint"].ToString()) == true) dicTemp.Add("Print", "打印");
                            if (bool.Parse(dt_power.Rows[0]["hasAudit"].ToString()) == true) dicTemp.Add("Auditing", "审核");

                            //有权限，CheckBox就勾选
                            string queryPostPower = "select hasQuery,hasInsert,hasUpdate,hasDelete,hasPrint,hasAudit from QMS_PermissionMenu where post='" + post + "'and model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"].ToString() + "'and menuName='" + tempNode3.Text + "'";
                            DataTable PostPower = DbAccess.SelectBySql(queryPostPower).Tables[0];

                            foreach (var item in dicTemp)
                            {
                                TreeNode tempNode4 = new TreeNode();
                                tempNode4.Name = item.Key.ToString();
                                tempNode4.Text = item.Value.ToString();

                                treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Nodes.Add(tempNode4);
                                treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Expand();

                                //有权限，CheckBox就勾选
                                if (tempNode4.Name == "Select") tempNode4.Checked = bool.Parse(PostPower.Rows[0]["hasQuery"].ToString());
                                else if (tempNode4.Name == "Insert") tempNode4.Checked = bool.Parse(PostPower.Rows[0]["hasInsert"].ToString());
                                else if (tempNode4.Name == "Update") tempNode4.Checked = bool.Parse(PostPower.Rows[0]["hasUpdate"].ToString());
                                else if (tempNode4.Name == "Delete") tempNode4.Checked = bool.Parse(PostPower.Rows[0]["hasDelete"].ToString());
                                else if (tempNode4.Name == "Print") tempNode4.Checked = bool.Parse(PostPower.Rows[0]["hasPrint"].ToString());
                                else if (tempNode4.Name == "Auditing") tempNode4.Checked = bool.Parse(PostPower.Rows[0]["hasAudit"].ToString());
                            }
                        }
                    }
                    treeView1.Nodes[0].Expand();

                }

            }

            treeView1.EndUpdate();

        }
        private void genEmputyTree()
        {
            treeView1.Nodes[0].Nodes.Clear();
            treeView1.BeginUpdate();
            //查询该管理员所能支配的权限
            //string queryHasPower = "select * from QMS_PermissionMenu where post='" + Login.manager + "'and permissions='是'";
            //DataTable dt0 = DbAccess.SelectBySql(queryHasPower).Tables[0];

            //查询该管理员所能支配的权限
            string parentNode = "select distinct model from QMS_PermissionMenu where post='" + Login.manager + "'and permissions='是'";
            DataTable dt0 = DbAccess.SelectBySql(parentNode).Tables[0];
            // dataGridView1.DataSource = dt0;
            if (dt0 != null && dt0.Rows.Count > 0)
            {
                foreach (DataRow dr in dt0.Rows)
                {
                    TreeNode tempNode1 = new TreeNode();
                    tempNode1.Name = dr["model"].ToString();
                    tempNode1.Text = dr["model"].ToString();
                    treeView1.Nodes[0].Nodes.Add(tempNode1);
                    tempNode1.ForeColor = Color.Blue;

                    string node2 = "select distinct modelGroup from QMS_PermissionMenu where model='" + dr["model"].ToString() + "'and  post='" + Login.manager + "'and permissions='是'";
                    DataTable md = DbAccess.SelectBySql(node2).Tables[0];

                    foreach (DataRow dr2 in md.Rows)
                    {
                        TreeNode tempNode2 = new TreeNode();
                        tempNode2.Name = dr2["modelGroup"].ToString();
                        tempNode2.Text = dr2["modelGroup"].ToString();

                        treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes.Add(tempNode2);
                        treeView1.Nodes[0].Nodes[dr["model"].ToString()].Expand();


                        string node3 = "select menuName from QMS_PermissionMenu where model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"] + "'and  post='" + Login.manager + "'and permissions='是'";
                        DataTable dt3 = DbAccess.SelectBySql(node3).Tables[0];


                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            TreeNode tempNode3 = new TreeNode();
                            tempNode3.Name = dr3["menuName"].ToString();
                            tempNode3.Text = dr3["menuName"].ToString();

                            treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes.Add(tempNode3);
                            treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Expand();


                            //无权限的项被disable
                            Dictionary<string, string> dicTemp = new Dictionary<string, string>();

                            string node4_queryPower = "select hasQuery,hasInsert,hasUpdate,hasDelete,hasPrint,hasAudit from QMS_PermissionMenu where model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"].ToString() + "'and menuName='" + dr3["menuName"].ToString() + "'and  post='" + Login.manager + "'and permissions='是'";
                            DataTable dt4 = DbAccess.SelectBySql(node4_queryPower).Tables[0];

                            if (bool.Parse(dt4.Rows[0]["hasQuery"].ToString()) == true) dicTemp.Add("Select", "查看");
                            if (bool.Parse(dt4.Rows[0]["hasInsert"].ToString()) == true) dicTemp.Add("Insert", "增加");
                            if (bool.Parse(dt4.Rows[0]["hasUpdate"].ToString()) == true) dicTemp.Add("Update", "修改");
                            if (bool.Parse(dt4.Rows[0]["hasDelete"].ToString()) == true) dicTemp.Add("Delete", "删除");
                            if (bool.Parse(dt4.Rows[0]["hasPrint"].ToString()) == true) dicTemp.Add("Print", "打印");
                            if (bool.Parse(dt4.Rows[0]["hasAudit"].ToString()) == true) dicTemp.Add("Auditing", "审核");

                            foreach (var item in dicTemp)
                            {

                                TreeNode node4 = new TreeNode();
                                node4.Name = item.Key.ToString();
                                node4.Text = item.Value.ToString();
                                treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Nodes.Add(node4);
                                treeView1.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Expand();
                            }


                        }
                    }
                    treeView1.Nodes[0].Expand();

                }

            }

            treeView1.EndUpdate();
        }

        private void txtpost_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtpost.Text != "")
            {
                string hasPower = "select * from QMS_PermissionMenu where post='" + txtpost.Text + "'";
                DataTable table = DbAccess.SelectBySql(hasPower).Tables[0];
                if (table != null && table.Rows.Count > 0)
                {
                    //genEmputyTree();
                    genTree(txtpost.Text);   //读取已有的权限
                }

                else
                    genEmputyTree();
            }
        }

        private void sBtnsave_Click(object sender, EventArgs e)
        {
            //保存岗位权限信息
            foreach (TreeNode node1 in treeView1.Nodes[0].Nodes)
            {
                if (node1.GetNodeCount(false) > 0)
                {
                    foreach (TreeNode node2 in node1.Nodes)
                    {
                        if (node2.GetNodeCount(false) > 0)
                        {
                            foreach (TreeNode node3 in node2.Nodes)
                            {
                                if (node3.GetNodeCount(false) > 0)
                                {
                                    bool query = false;
                                    bool insert = false;
                                    bool update = false;
                                    bool delete = false;
                                    bool print = false;
                                    bool audit = false;
                                    string permission = "否";
                                    foreach (TreeNode node4 in node3.Nodes)
                                    {
                                        if (node4.Name == "Select") query = node4.Checked;
                                        if (node4.Name == "Insert") insert = node4.Checked;
                                        if (node4.Name == "Update") update = node4.Checked;
                                        if (node4.Name == "Delete") delete = node4.Checked;
                                        if (node4.Name == "Print") print = node4.Checked;
                                        if (node4.Name == "Auditing") audit = node4.Checked;

                                    }
                                    if (query || insert || update || delete || print || audit)
                                    {
                                        permission = "是";
                                    }
                                    string check = "select * from QMS_PermissionMenu where model='" + node1.Text + "'and modelGroup='" + node2.Text + "'and MenuName='" + node3.Text + "'and post='" + txtpost.Text + "'";
                                    DataTable table = DbAccess.SelectBySql(check).Tables[0];
                                    if (table != null && table.Rows.Count > 0)
                                    {
                                        //更新
                                        string sql_update = "update QMS_PermissionMenu set permissions='" + permission + "', hasQuery='" + query + "',hasInsert='" + insert + "',hasUpdate='" + update + "',hasDelete='" + delete + "',hasPrint='" + print + "',hasAudit='" + audit + "',updateUser='" + Login.username + "',updateTime='" + DateTime.Now + "'"
                                        + "where model='" + node1.Text + "'and modelGroup='" + node2.Text + "'and MenuName='" + node3.Text + "'and post='" + txtpost.Text + "'";
                                        sql_update = sql_update.Replace("True", "1");
                                        sql_update = sql_update.Replace("False", "0");
                                        DbAccess.ExecuteSql(sql_update);
                                    }
                                    else
                                    {
                                        //新增
                                        string sql = "insert into QMS_PermissionMenu(model,modelGroup,MenuName,post,permissions,hasQuery,hasInsert,hasUpdate,hasDelete,hasPrint,hasAudit,updateUser,updateTime)"
                                               + "values('" + node1.Text + "','" + node2.Text + "','" + node3.Text + "','" + txtpost.Text + "','" + permission + "','" + query + "','" + insert + "','" + update + "','" + delete + "','" + print + "','" + audit + "','" + Login.username + "','" + DateTime.Now + "')";

                                        sql = sql.Replace("True", "1");
                                        sql = sql.Replace("False", "0");
                                        DbAccess.ExecuteSql(sql);
                                    }

                                }

                            }

                        }
                    }
                }
            }
            MessageBox.Show("保存成功！");
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Nodes.Count > 0)
            {
                foreach (TreeNode node in e.Node.Nodes)
                {
                    node.Checked = e.Node.Checked;
                }
            }
        }
    }
}