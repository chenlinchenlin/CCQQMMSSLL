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
    public partial class QMSassignManagerPower : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public QMSassignManagerPower()
        {
            InitializeComponent();
            bindData();
        }

        private void bindData()
        {
            txtgroupname.Properties.Items.Clear();
            //string sql = @"select distinct manager from QMS_UserInfo where manager<>'IT管理员'";
            //DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            //if (dt != null && dt.Rows.Count > 0)
            //{
            //    foreach (DataRow dr in dt.Rows)
            //    {
            //        comboBox3.Items.Add(dr["manager"]);
            //    }
            //}
            ////comboBox3.Items.Remove("IT管理员");

            txtgroupname.Properties.Items.Add("QE管理员");
            txtgroupname.Properties.Items.Add("SQE管理员");
            txtgroupname.Properties.Items.Add("IQC管理员");
            txtgroupname.Properties.Items.Add("OQC管理员");
            txtgroupname.Properties.Items.Add("IPQC管理员");
            txtgroupname.Properties.Items.Add("IPQC&OQC管理员");
            txtgroupname.Properties.Items.Add("ESD管理员");
            txtgroupname.Properties.Items.Add("外访组管理员");
            txtgroupname.Properties.Items.Add("IT管理员");
            //comboBox3.Items.Add("主管以上管理员");
        }
        private void txtuserid_Leave(object sender, EventArgs e)
        {
            string sql = @"select userName,post from QMS_userInfo where userId='" + txtuserid.Text + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                txtusename.Text = dt.Rows[0]["userName"].ToString();
                txtuserpost.Text = dt.Rows[0]["post"].ToString();
            }
            else
            {
                MessageBox.Show("用户名不存在");
                return;
            }
        }

        private void txtuserid_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyValue == 13 && txtuserid.Text != "")
            {
                string sql = @"select userName,post from QMS_userInfo where userId='" + txtuserid.Text + "'";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

                if (dt != null && dt.Rows.Count > 0)
                {
                    txtusename.Text = dt.Rows[0]["userName"].ToString();
                    txtuserpost.Text = dt.Rows[0]["post"].ToString();
                }
                else
                {
                    MessageBox.Show("用户名不存在");
                    return;
                }
            }

        }


        private void genTree(string post)
        {
            treeView.Nodes[0].Nodes.Clear();
            treeView.BeginUpdate();

            string parentNode = "select distinct model from QMS_PermissionMenu where post='" + post + "'";
            DataTable dt0 = DbAccess.SelectBySql(parentNode).Tables[0];

            if (dt0 != null && dt0.Rows.Count > 0)
            {
                Dictionary<string, string> dicTemp = new Dictionary<string, string>();
                //string strFunc = "select fID,fName from Functions";
                //DataTable dtFunc = DbAccess.SelectBySql(strFunc).Tables[0];

                //foreach (DataRow dtF in dtFunc.Rows)
                //{
                //    dicTemp.Add(dtF["fID"].ToString(), dtF["fName"].ToString());
                //}
                dicTemp.Add("Select", "查看");
                dicTemp.Add("Insert", "增加");
                dicTemp.Add("Update", "修改");
                dicTemp.Add("Delete", "删除");
                dicTemp.Add("Print", "打印");
                dicTemp.Add("Auditing", "审核");

                foreach (DataRow dr in dt0.Rows)
                {
                    TreeNode tempNode1 = new TreeNode();
                    tempNode1.Name = dr["model"].ToString();
                    tempNode1.Text = dr["model"].ToString();
                    treeView.Nodes[0].Nodes.Add(tempNode1);
                    tempNode1.ForeColor = Color.Blue;

                    string node2 = "select distinct modelGroup from QMS_PermissionMenu where  post='" + post + "'and model='" + dr["model"].ToString() + "'";
                    DataTable md = DbAccess.SelectBySql(node2).Tables[0];

                    foreach (DataRow dr2 in md.Rows)
                    {
                        TreeNode tempNode2 = new TreeNode();
                        tempNode2.Name = dr2["modelGroup"].ToString();
                        tempNode2.Text = dr2["modelGroup"].ToString();

                        treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes.Add(tempNode2);
                        treeView.Nodes[0].Nodes[dr["model"].ToString()].Expand();


                        string node3 = "select menuName from QMS_PermissionMenu where  post='" + post + "'and model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"].ToString() + "'";
                        DataTable dt3 = DbAccess.SelectBySql(node3).Tables[0];


                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            TreeNode tempNode3 = new TreeNode();
                            tempNode3.Name = dr3["menuName"].ToString();
                            tempNode3.Text = dr3["menuName"].ToString();

                            treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes.Add(tempNode3);
                            treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Expand();

                            string queryPower = "select hasQuery,hasInsert,hasUpdate,hasDelete,hasPrint,hasAudit from QMS_PermissionMenu where post='" + post + "'and model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"].ToString() + "'and menuName='" + tempNode3.Text + "'";
                            DataTable dt_power = DbAccess.SelectBySql(queryPower).Tables[0];


                            foreach (var item in dicTemp)
                            {
                                TreeNode tempNode4 = new TreeNode();
                                tempNode4.Name = item.Key.ToString();
                                tempNode4.Text = item.Value.ToString();

                                treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Nodes.Add(tempNode4);
                                treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Expand();

                                //有权限，CheckBox就勾选

                                if (tempNode4.Name == "Select") tempNode4.Checked = bool.Parse(dt_power.Rows[0]["hasQuery"].ToString());
                                if (tempNode4.Name == "Insert") tempNode4.Checked = bool.Parse(dt_power.Rows[0]["hasInsert"].ToString()); ;
                                if (tempNode4.Name == "Update") tempNode4.Checked = bool.Parse(dt_power.Rows[0]["hasUpdate"].ToString());
                                if (tempNode4.Name == "Delete") tempNode4.Checked = bool.Parse(dt_power.Rows[0]["hasDelete"].ToString());
                                if (tempNode4.Name == "Print") tempNode4.Checked = bool.Parse(dt_power.Rows[0]["hasPrint"].ToString());
                                if (tempNode4.Name == "Auditing") tempNode4.Checked = bool.Parse(dt_power.Rows[0]["hasAudit"].ToString());
                            }
                        }
                    }
                    treeView.Nodes[0].Expand();

                }

            }

            treeView.EndUpdate();
        }

        private void genEmputyTree()
        {
            treeView.Nodes[0].Nodes.Clear();
            treeView.BeginUpdate();

            string parentNode = "select distinct model from QMS_AutoExtractMenu";
            DataTable dt0 = DbAccess.SelectBySql(parentNode).Tables[0];

            if (dt0 != null && dt0.Rows.Count > 0)
            {
                Dictionary<string, string> dicTemp = new Dictionary<string, string>();
                //string strFunc = "select fID,fName from Functions";
                //DataTable dtFunc = DbAccess.SelectBySql(strFunc).Tables[0];

                //foreach (DataRow dtF in dtFunc.Rows)
                //{
                //    dicTemp.Add(dtF["fID"].ToString(), dtF["fName"].ToString());
                //}

                dicTemp.Add("Select", "查看");
                dicTemp.Add("Insert", "增加");
                dicTemp.Add("Update", "修改");
                dicTemp.Add("Delete", "删除");
                dicTemp.Add("Print", "打印");
                dicTemp.Add("Auditing", "审核");

                foreach (DataRow dr in dt0.Rows)
                {
                    TreeNode tempNode1 = new TreeNode();
                    tempNode1.Name = dr["model"].ToString();
                    tempNode1.Text = dr["model"].ToString();
                    treeView.Nodes[0].Nodes.Add(tempNode1);
                    tempNode1.ForeColor = Color.Blue;

                    string node2 = "select distinct modelGroup from QMS_AutoExtractMenu where model='" + dr["model"].ToString() + "'";
                    DataTable md = DbAccess.SelectBySql(node2).Tables[0];

                    foreach (DataRow dr2 in md.Rows)
                    {
                        TreeNode tempNode2 = new TreeNode();
                        tempNode2.Name = dr2["modelGroup"].ToString();
                        tempNode2.Text = dr2["modelGroup"].ToString();

                        treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes.Add(tempNode2);
                        treeView.Nodes[0].Nodes[dr["model"].ToString()].Expand();


                        string node3 = "select menuName from QMS_AutoExtractMenu where model='" + dr["model"].ToString() + "'and modelGroup='" + dr2["modelGroup"] + "'";
                        DataTable dt3 = DbAccess.SelectBySql(node3).Tables[0];


                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            TreeNode tempNode3 = new TreeNode();
                            tempNode3.Name = dr3["menuName"].ToString();
                            tempNode3.Text = dr3["menuName"].ToString();

                            treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes.Add(tempNode3);
                            treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Expand();


                            foreach (var item in dicTemp)
                            {
                                TreeNode tempNode4 = new TreeNode();
                                tempNode4.Name = item.Key.ToString();
                                tempNode4.Text = item.Value.ToString();

                                treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Nodes.Add(tempNode4);
                                treeView.Nodes[0].Nodes[dr["model"].ToString()].Nodes[dr2["modelGroup"].ToString()].Nodes[dr3["menuName"].ToString()].Expand();
                            }
                        }
                    }
                    treeView.Nodes[0].Expand();

                }

            }

            treeView.EndUpdate();
        }
        private void txtgroupname_SelectedValueChanged(object sender, EventArgs e)
        {
            if (txtgroupname.Text == "" || txtuserpost.Text == "")
                return;
            else
            {
                string manager = "select * from QMS_PermissionMenu where post='" + txtgroupname.Text + "'";
                DataTable dt = DbAccess.SelectBySql(manager).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    genEmputyTree();
                    ////genTree(txtgroupname.Text);//读取已有权限的树形结构
                }
                else
                {
                    genEmputyTree();
                }
            }
        }

        private void sBtnsave_Click(object sender, EventArgs e)
        {
            //判断不能重复保存
            /*  未完成  */




            //保存管理员信息
            if (txtgroupname.Text == "")
            {
                MessageBox.Show("请填写权组信息");
                return;
            }
            string strSql = "update QMS_userInfo set manager='" + txtgroupname.Text + "',updateUser='" + Login.username + "',updateTime='" + DateTime.Now + "'where userId='" + txtuserid.Text + "'";
            DbAccess.ExecuteSql(strSql);

            //保存管理员的权限信息,分新增和更新
            foreach (TreeNode node1 in treeView.Nodes[0].Nodes)
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
                                    //string function_Sql = null;

                                    bool query = false;
                                    bool insert = false;
                                    bool update = false;
                                    bool delete = false;
                                    bool print = false;
                                    bool audit = false;
                                    string permission = "否";
                                    foreach (TreeNode node4 in node3.Nodes)
                                    {
                                        //function_Sql = function_Sql + "has" + node4.Name + "= '" + node4.Checked + "',";
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

                                    //判断是新增还是更新
                                    string check = "select * from QMS_PermissionMenu where model='" + node1.Text + "'and modelGroup='" + node2.Text + "'and MenuName='" + node3.Text + "'and post='" + txtgroupname.Text + "'";
                                    DataTable table = DbAccess.SelectBySql(check).Tables[0];
                                    if (table != null && table.Rows.Count > 0)
                                    {
                                        //更新
                                        string sql_update = "update QMS_PermissionMenu set permissions='" + permission + "', hasQuery='" + query + "',hasInsert='" + insert + "',hasUpdate='" + update + "',hasDelete='" + delete + "',hasPrint='" + print + "',hasAudit='" + audit + "',updateUser='" + Login.username + "',updateTime='" + DateTime.Now + "'"
                                                            + "where model='" + node1.Text + "'and modelGroup='" + node2.Text + "'and MenuName='" + node3.Text + "'and post='" + txtgroupname.Text + "'";
                                        sql_update = sql_update.Replace("True", "1");
                                        sql_update = sql_update.Replace("False", "0");
                                        DbAccess.ExecuteSql(sql_update);
                                    }

                                    else
                                    {
                                        //新增
                                        string sql = "insert into QMS_PermissionMenu(model,modelGroup,MenuName,post,permissions,hasQuery,hasInsert,hasUpdate,hasDelete,hasPrint,hasAudit,updateUser,updateTime)"
                                                    + "values('" + node1.Text + "','" + node2.Text + "','" + node3.Text + "','" + txtgroupname.Text + "','" + permission + "','" + query + "','" + insert + "','" + update + "','" + delete + "','" + print + "','" + audit + "','" + Login.username + "','" + DateTime.Now + "')";

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

        private void treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Nodes.Count > 0)
            {
                foreach (TreeNode node in e.Node.Nodes)
                {
                    node.Checked = e.Node.Checked;
                }
            }
        }

        private void treeView_AfterCheck(object sender, TreeViewEventArgs e)
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