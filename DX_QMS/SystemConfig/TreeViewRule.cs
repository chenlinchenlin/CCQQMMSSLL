using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Collections;
using DX_QMS.Common;

namespace DX_QMS.SystemConfig
{
    public partial class TreeViewRule : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public TreeViewRule()
        {
            InitializeComponent();
            setRule();
        }
        private void setRule()
        {
            Dictionary<string, bool> dic = GroupPermission.SelectRulesForForm("TreeViewRule", Login.groupId);
            btnSave.Enabled = dic["hasUpdate"];
        }

        private void TreeViewRule_Load(object sender, EventArgs e)
        {
            if (Login.deptId != "IT")
                lblIsWeb.Visible = cbIsWeb.Visible = false;
            cbIsWeb.SelectedIndex = 0;
        }
        protected void bindGroup()
        {
            DataSet ds = Groups.SelectAll(Login.groupId, Login.deptId, cbIsWeb.Text);
            if (ds.Tables.Count > 0)
            {
                cbGroupID.DataSource = ds.Tables[0];
                cbGroupID.DisplayMember = ds.Tables[0].Columns["groupname"].ToString();
                cbGroupID.ValueMember = ds.Tables[0].Columns["groupid"].ToString();
                cbGroupID.SelectedIndex = -1;
            }
        }
        private void cbIsWeb_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbGroupID.SelectedIndexChanged -= new EventHandler(cbGroupID_SelectedIndexChanged);
            bindGroup();
            treeRule.Nodes[0].Nodes.Clear();
            cbGroupID.SelectedIndexChanged += new EventHandler(cbGroupID_SelectedIndexChanged);
        }
        protected void genTree(string IsWeb)
        {
            treeRule.Nodes[0].Nodes.Clear();
            treeRule.BeginUpdate();

            string strSql = "select * from Module where IsWeb='" + IsWeb + "' and IsOper='Y' and mID <>'SystemHelp' order by mNo ";
            DataTable dt = DbAccess.SelectBySql(strSql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    TreeNode tempNode1 = new TreeNode();
                    tempNode1.Name = dr["mID"].ToString();
                    tempNode1.Text = dr["mName"].ToString();
                    treeRule.Nodes[0].Nodes.Add(tempNode1);
                    tempNode1.ForeColor = Color.Blue;
                }
                treeRule.Nodes[0].Expand();
                string strSql1 = " select m.mID, m.mName, m.mWindow, m.mNo, m.mParent, m.mRule, m.IsOper from Module AS m";
                strSql1 += " INNER JOIN Groups AS g ON m.IsWeb = g.IsWeb AND CHARINDEX(g.deptid, m.mDepts) > 0";
                strSql1 += " where m.IsWeb='" + IsWeb + "' and m.IsOper='N' and m.mParent <>'SystemHelp' and g.groupid ='" + cbGroupID.SelectedValue.ToString() + "' order by mNo";
                DataTable dt1 = DbAccess.SelectBySql(strSql1).Tables[0];
                if (dt1.Rows.Count > 0)
                {
                    Dictionary<string, string> dicTemp = new Dictionary<string, string>();

                    string strFunc = "select * from Functions";
                    DataTable dtFunc = DbAccess.SelectBySql(strFunc).Tables[0];
                    for (int i = 0; i < dtFunc.Rows.Count; i++)
                    {
                        dicTemp.Add(dtFunc.Rows[i][0].ToString(), dtFunc.Rows[i][1].ToString());
                    }

                    foreach (DataRow dr in dt1.Rows)
                    {
                        TreeNode tempNode2 = new TreeNode();
                        tempNode2.Name = dr["mID"].ToString();
                        tempNode2.Text = dr["mName"].ToString();
                        treeRule.Nodes[0].Nodes[dr["mParent"].ToString()].Nodes.Add(tempNode2);
                        treeRule.Nodes[0].Nodes[dr["mParent"].ToString()].Expand();

                        if (dr["mRule"].ToString() != "")
                        {
                            string[] temp = dr["mRule"].ToString().Split(',');
                            for (int i = 0; i < temp.Length; i++)
                            {
                                TreeNode newNode = new TreeNode();
                                newNode.Name = temp[i].ToString();
                                newNode.Text = dicTemp[temp[i].ToString()].ToString();
                                tempNode2.Nodes.Add(newNode);
                            }
                        }
                    }
                }
            }
            treeRule.EndUpdate();
        }
        protected bool hasGroupID(string groupId)
        {
            bool flag = false;
            string str = "select moduleID from GroupPermission where groupID='" + groupId + "'";
            DataSet ds = DbAccess.SelectBySql(str);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                flag = true;
            return flag;
        }
        protected void bindTree(string groupID)
        {
            treeRule.AfterCheck -= treeRule_AfterCheck;
            string str = "select A.*,B.mParent,B.mDepts from GroupPermission A left join Module B ON A.moduleID=B.mID";
            str += " INNER JOIN Groups AS g ON g.groupid = A.groupID AND CHARINDEX(g.deptid, B.mDepts) > 0 ";
            str += " where  A.groupID='" + groupID + "'";
            DataTable dt = DbAccess.SelectBySql(str).Tables[0];
            if (dt == null || dt.Rows.Count < 1) return;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                foreach (TreeNode ndModule in treeRule.Nodes[0].Nodes[dt.Rows[i]["mParent"].ToString()].Nodes)
                {
                    if (ndModule.Name == dt.Rows[i]["moduleID"].ToString())
                    {
                        bool bIsAllChecked = true;
                        foreach (TreeNode node in ndModule.Nodes)
                        {
                            node.Checked = bool.Parse(dt.Rows[i]["has" + node.Name].ToString());
                            bIsAllChecked = bIsAllChecked & node.Checked;
                        }
                        ndModule.Checked = bIsAllChecked;
                    }
                }
            }
            treeRule.AfterCheck += treeRule_AfterCheck;
        }
        private void cbGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            genTree(cbIsWeb.Text);
            if (hasGroupID(cbGroupID.SelectedValue.ToString()))
                bindTree(cbGroupID.SelectedValue.ToString());
            else
                treeRule.Nodes[0].Checked = false;
            treeRule.Select();
        }
        protected void insertNewModule()
        {
            DataSet groupModule = GroupPermission.SelectModuleByGroupID(cbGroupID.SelectedValue.ToString());
            if (groupModule.Tables.Count < 1)
            {
                GroupPermission.AddNewRecordByGroupID(cbGroupID.SelectedValue.ToString(), cbIsWeb.Text);
                return;
            }
            DataSet exitModule = Module.SelectAllModuleID(cbIsWeb.Text);
            if (groupModule.Tables[0].Rows.Count == exitModule.Tables[0].Rows.Count)
                return;
            ArrayList moduleArr = new ArrayList();
            for (int i = 0; i < exitModule.Tables[0].Rows.Count; i++)
            {
                moduleArr.Add(exitModule.Tables[0].Rows[i][0].ToString());
            }

            for (int j = 0; j < groupModule.Tables[0].Rows.Count; j++)
            {
                if (moduleArr.Contains(groupModule.Tables[0].Rows[j][0].ToString()))
                    moduleArr.Remove(groupModule.Tables[0].Rows[j][0].ToString());
            }

            string str = "";
            ArrayList strList = new ArrayList();

            for (int z = 0; z < moduleArr.Count; z++)
            {
                str = "insert into GroupPermission (groupID,moduleID) values('" + cbGroupID.SelectedValue.ToString() + "','" + moduleArr[z].ToString() + "')";
                strList.Add(str);
            }
            DbAccess.ExecutSqlTran(strList);
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            insertNewModule();
            ArrayList arr1 = new ArrayList();
            foreach (TreeNode node1 in treeRule.Nodes[0].Nodes)
            {
                if (node1.GetNodeCount(false) > 0)
                {
                    foreach (TreeNode node2 in node1.Nodes)
                    {
                        if (node2.GetNodeCount(false) > 0)
                        {
                            string strSql = "update GroupPermission set ";
                            string whereSql = "updateUser='" + Login.userId + "', updateTime='" + DateTime.Now + "' where ";

                            foreach (TreeNode node3 in node2.Nodes)
                            {
                                strSql = strSql + "has" + node3.Name + " =" + node3.Checked + ", ";
                            }
                            whereSql = whereSql + "groupID='" + cbGroupID.SelectedValue.ToString() + "' and moduleID= '" + node2.Name + "'";
                            strSql = strSql + whereSql;
                            strSql = strSql.Replace("True", "1");
                            strSql = strSql.Replace("False", "0");
                            arr1.Add(strSql);
                        }
                    }
                }
            }

            if (DbAccess.ExecutSqlTran(arr1))
                MessageBox.Show("保存成功！");
            bindTree(cbGroupID.SelectedValue.ToString());
        }

        private void treeRule_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Nodes.Count > 0)
            {
                foreach (TreeNode node in e.Node.Nodes)
                {
                    node.Checked = e.Node.Checked;
                }
            }
        }

        private void treeRule_AfterExpand(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Nodes.Count > 0)
            {
                treeRule.BeginUpdate();
                e.Node.ExpandAll();
                treeRule.EndUpdate();
            }
        }
    }
}