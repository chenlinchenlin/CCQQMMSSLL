using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;
using System.IO;
using DevExpress.XtraEditors.Controls;
using System.Collections;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using System.Diagnostics;

namespace DX_QMS.IPQC
{
    public partial class IPQCExceptionList : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "192.168.0.204\\FilePath$";
        public IPQCExceptionList()
        {
            InitializeComponent();
        }


        private void bindorg_id()
        {
            string sql = "select ORG_ID,ORG_NAME from ORG_INFO where ORG_TYPE='ORG' order by SORT ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtorg_id.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtorg_id.Properties.Items.Add(row["ORG_NAME"]);
            }
            txtorg_id.SelectedIndex = 1;

        }
        private void bindstandid()
        {
            string sql = "	select distinct Progsetvalue from  IPQCProgset where Progsettype = '站别' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtstandid.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtstandid.Properties.Items.Add(row["Progsetvalue"]);
            }
            txtstandid.SelectedIndex = 0;
        }
        private void bindproductlineid()
        {
            string sql = "	select distinct Progsetvalue from  IPQCProgset where Progsettype = '生产线别' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtproductlineid.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtproductlineid.Properties.Items.Add(row["Progsetvalue"]);
            }
            txtproductlineid.SelectedIndex = 0;
        }
        private void bindbadclass()
        {
            string sql = "	select distinct Progsetvalue from  IPQCProgset where Progsettype = '不良类别' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtbadclass.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtbadclass.Properties.Items.Add(row["Progsetvalue"]);
            }
            txtbadclass.SelectedIndex = 0;
        }
        private void binddutyDepartment()
        {
            string sql = "	select distinct Progsetvalue from  IPQCProgset where Progsettype = '责任部门' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtdutyDepartment.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtdutyDepartment.Properties.Items.Add(row["Progsetvalue"]);
            }
            txtdutyDepartment.SelectedIndex = 0;
        }
        private void bindchargeMan()
        {
            string sql = "	select distinct Progsetvalue from  IPQCProgset where Progsettype = '责任人' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtchargeMan.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtchargeMan.Properties.Items.Add(row["Progsetvalue"]);
            }
            txtchargeMan.SelectedIndex = 0;
        }
        private void bindoverdutyDepartment()
        {
            string sql = "	select distinct Progsetvalue from  IPQCProgset where Progsettype = '延时责任部门' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtoverdutyDepartment.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtoverdutyDepartment.Properties.Items.Add(row["Progsetvalue"]);
            }
            txtoverdutyDepartment.SelectedIndex = 0;
        }

        private void setRule()
        {
            string post = "";
            if (!string.IsNullOrEmpty(Login.manager))
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "制程异常");
            this.sBtnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.sBtndelete.Enabled = bool.Parse(dic["hasDelete"].ToString());
            this.sBtnupdate.Enabled = bool.Parse(dic["hasUpdate"].ToString());
        }

        private void IPQCExceptionList_Load(object sender, EventArgs e)
        {
            bindorg_id();
            bindstandid();
            bindproductlineid();
            bindbadclass();
            binddutyDepartment();
            bindchargeMan();
            bindoverdutyDepartment();
            ////  setRule();
            txtmodel.ReadOnly = false;
            txtcustomer.ReadOnly = false;
            txtworknoqty.ReadOnly = true;
            txtfloor.SelectedIndex = 0;
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
            resetcontrol();
            gridControl.DataSource = null;
        }
        private void txtworkno_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtworkno.Text != "")
            {
                txtworkno_Leave(sender, e);
            }
        }
        private void txtworkno_Leave(object sender, EventArgs e)
        {
            gridControl.DataSource = null;
            if (txtworkno.Text == "")
                return;
            txtworkno.Leave -= txtworkno_Leave;
            string sql = "select item_num,item_desc,start_quantity,job_desc,WO.DEPARTMENT_ID,completion_subinventory,order_number,gn_mo,V.ORGANIZATION_ID from cux_wip_header_v v,CUX_WIP_OPERATIONS_V WO  where " +
                                " WO.WIP_ENTITY_ID =v.WIP_ENTITY_ID and wip_entity_name =" + "'" + txtworkno.Text.Trim().ToUpper() + "' and v.ORGANIZATION_CODE = '" + txtorg_id.Text + "'";
            DataSet ds = DbAccess.SelectByOracle(sql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                txthytcode.Text = ds.Tables[0].Rows[0]["item_num"].ToString();
                txtworknoqty.Text = ds.Tables[0].Rows[0]["start_quantity"].ToString();
                ArrayList ltL = new ArrayList();
                ArrayList customerL = new ArrayList();
                int index = 0, customerindex = 0;
                string str = ds.Tables[0].Rows[0]["item_desc"].ToString();
                foreach (Char ch in str)
                {
                    if (ch == '(' || ch == '（')
                    {
                        ltL.Add(index);
                    }
                    if (ch == '【')
                    {
                        customerL.Add(customerindex);
                    }
                    index++;
                    customerindex++;
                }
                ArrayList ltR = new ArrayList();
                ArrayList customerR = new ArrayList();
                int indexR = 0, customerindexR = 0;
                foreach (Char ch in str)
                {
                    if (ch == ')' || ch == '）')
                    {
                        ltR.Add(indexR);
                    }
                    if (ch == '】')
                    {
                        customerR.Add(customerindexR);
                    }
                    indexR++;
                    customerindexR++;
                }
                try
                {
                    try
                    {
                        txtmodel.Text = str.Substring(int.Parse(ltL[0].ToString()) + 1, int.Parse(ltR[0].ToString()) - int.Parse(ltL[0].ToString()) - 1);
                    }
                    catch
                    {
                        txtmodel.Text = "";
                    }

                    txtcustomer.Text = str.Substring(int.Parse(customerL[0].ToString()) + 1, int.Parse(customerR[0].ToString()) - int.Parse(customerL[0].ToString()) - 1).Replace("专用", "");

                    if (txtcustomer.Text.Contains("申请"))
                    {
                        txtcustomer.Text = str.Substring(int.Parse(customerL[0].ToString()) + 1, int.Parse(customerR[0].ToString()) - int.Parse(customerL[0].ToString()) - 1).Replace("申请", "");
                    }
                    if (txtcustomer.Text == "保税")
                    {
                        txtcustomer.Text = str.Substring(int.Parse(customerL[1].ToString()) + 1, int.Parse(customerR[1].ToString()) - int.Parse(customerL[1].ToString()) - 1).Replace("专用", "");
                    }
                    if (txtcustomer.Text == "返工")
                    {
                        txtcustomer.Text = str.Substring(int.Parse(customerL[1].ToString()) + 1, int.Parse(customerR[1].ToString()) - int.Parse(customerL[1].ToString()) - 1).Replace("专用", "");
                    }
                    if (txtcustomer.Text == "16004")
                    {
                        txtcustomer.Text = "新飞通";
                    }
           
                }              
                catch
                {
                   
                    if (txthytcode.Text.Trim().StartsWith("1150") || txthytcode.Text.Trim().StartsWith("6103") || txthytcode.Text.Trim().StartsWith("1151"))
                    {
                        txtcustomer.Text = "主营";
                        txtmodel.Text = str;
                    }
                }
                txtworkno.ReadOnly = true;               
                txtmodel.ReadOnly = true;
                txtcustomer.ReadOnly = true;
                txtworknoqty.ReadOnly = true;
                sBtnsave.Enabled = true;
            }
            else
            {
                MessageBox.Show(txtworkno.Text + "该工单号不存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtworkno.Text = "";
                txtworkno.Focus();
                return;
            }         
            txtworkno.Leave += txtworkno_Leave;
            selectdata();
        }

        private void txtqty_Leave(object sender, EventArgs e)
        {          
            int outputqty = 0;
            if (!int.TryParse(txtqty.Text.Trim()=="" ? "0" : txtqty.Text, out outputqty))
            {
                MessageBox.Show("投入数数量不正确", "提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
        }

        private void txtNGQty_Leave(object sender, EventArgs e)
        {
            int outputNGQty = 0;
            if (!int.TryParse(txtNGQty.Text.Trim() == "" ? "0" : txtqty.Text, out outputNGQty))
            {
                MessageBox.Show("不良数量不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void sBtnbadpicture_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "图像文件(*.jpg;bmp;png;jpeg)|*.jpg;*.bmp;*.png;*.jpeg";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;
            txtbadpicture.Text = ofdImport.FileName;
            FileStream fs = new FileStream(txtbadpicture.Text, FileMode.Open);
            Image bt = Image.FromStream(fs);
            fs.Close();
            fs.Dispose();
            picbadImage.Image = bt;
            picbadImage.Properties.SizeMode = PictureSizeMode.Stretch;
        }

        public void saveselect(string item)
        {       
            string sql = @"  select  ifrepeat 是否重复发生,standid 站别,item 表单编号,checkdate 日期,productlineid 生产线别,customer 客户,org_id 组织,
		        workno 工单号,worknoqty 工单数量,hytcode Hytera编码,model 客户机型,qty 投入数,NGQty 不良数,CASE WHEN NGQty < worknoqty  then  convert( varchar,convert(numeric(3,1),(NGQty+0.0)/(isnull(worknoqty,qty))*100 ))+'%' ELSE '100%' end 不良率,
				badclass 不良类别,baddescribe 问题描述,problemtype 问题分类,temporaryhandle 临时处理方法,CauseAnalysis 原因分析,improvemeasures 改善计划,
				dutyDepartment 责任部门,updateMan 记录人,chargeMan 责任人,ifClose 是否关闭,ifOvertime 是否超期,overdutyDepartment 延时责任部门,badImage 不良图片
		        from IPQCExceptionList where  item = '" + item+ "' order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;              
            }           
        }


        public string IPQCExceptioninsertList( string opertype,string item,string org_id, string workno,int worknoqty, string hytcode,string model,int qty,int NGQty,string ifrepeat,
                        string standid,string productlineid,string customer,string badclass,string baddescribe,string problemtype,string temporaryhandle,
                        string CauseAnalysis,string improvemeasures,string dutyDepartment,string kpiData, string updateMan, string chargeMan,string ifClose,string ifOvertime,string overdutyDepartment)
        {
            SqlParameter[] para = new SqlParameter[27];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@item", item);
            para[2] = new SqlParameter("@org_id", org_id);
            para[3] = new SqlParameter("@workno", workno);
            para[4] = new SqlParameter("@worknoqty", worknoqty);
            para[5] = new SqlParameter("@hytcode", hytcode);
            para[6] = new SqlParameter("@model", model);
            para[7] = new SqlParameter("@qty", qty);
            para[8] = new SqlParameter("@NGQty", NGQty);
            para[9] = new SqlParameter("@ifrepeat", ifrepeat);
            para[10] = new SqlParameter("@standid", standid);
            para[11] = new SqlParameter("@productlineid", productlineid);
            para[12] = new SqlParameter("@customer", customer);
            para[13] = new SqlParameter("@badclass", badclass);
            para[14] = new SqlParameter("@baddescribe", baddescribe);
            para[15] = new SqlParameter("@problemtype", problemtype);
            para[16] = new SqlParameter("@temporaryhandle", temporaryhandle);
            para[17] = new SqlParameter("@CauseAnalysis", CauseAnalysis);
            para[18] = new SqlParameter("@improvemeasures", improvemeasures);
            para[19] = new SqlParameter("@dutyDepartment", dutyDepartment);
            para[20] = new SqlParameter("@kpiData", kpiData);
            para[21] = new SqlParameter("@updateMan", updateMan);
            para[22] = new SqlParameter("@chargeMan", chargeMan);
            para[23] = new SqlParameter("@ifClose", ifClose);
            para[24] = new SqlParameter("@ifOvertime",ifOvertime);
            para[25] = new SqlParameter("@overdutyDepartment", overdutyDepartment);
            para[26] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[26].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IPQCExceptioninsertList", para);
            return para[26].Value.ToString();
        }

        private void sBtnsave_Click(object sender, EventArgs e)
        {
            if (txtworkno.Text.Trim() == "")
                return;
            if (txtqty.Text.Trim() == "" || txtNGQty.Text.Trim() == "")
            {
                MessageBox.Show("请输入投入数量和不良数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int worknoqty = 0;
            if (!int.TryParse(txtworknoqty.Text.Trim() == "" ? "0" : txtworknoqty.Text, out worknoqty))
            {
                MessageBox.Show("工单数量不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (worknoqty <1)
            {
                MessageBox.Show("工单数量不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int outputqty = 0;
            if (!int.TryParse(txtqty.Text.Trim() == "" ? "0" : txtqty.Text, out outputqty))
            {
                MessageBox.Show("投入数数量不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int outputNGQty = 0;
            if (!int.TryParse(txtNGQty.Text.Trim() == "" ? "0" : txtqty.Text, out outputNGQty))
            {
                MessageBox.Show("不良数量不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (int.Parse(txtNGQty.Text) > int.Parse(txtqty.Text))
            {
                MessageBox.Show("不良数量不能大于投入数量","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }





            ////string item = DateTime.Now.ToString("yyyyMMddHHmmss") +"-"+txtfloor.Text.Trim();

            //int itemmath = 0;
            //if (!int.TryParse(txtitem.Text.Trim() == "" ? "0" : txtitem.Text, out itemmath))
            //{
            //    MessageBox.Show("表单序列号不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    txtitem.Text = "";
            //    return;
            //}

            string item = DateTime.Now.ToString("yyyyMMdd") + "-" + txtfloor.Text.Trim();
            // string checkitem = item + txtitem.Text.Trim();

            //string sqlcheckitem = @" select 1 from IPQCExceptionList  where item = '"+ item + "' ";
            //DataTable dt = DbAccess.SelectBySql(sqlcheckitem).Tables[0];
            //if (dt != null && dt.Rows.Count > 0)
            //{
            //    MessageBox.Show("该表单编号已存在","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
            //    return;
            //}

            //string sqlcheckitem = @"  select max(item) item  from IPQCExceptionList where item like  '%"+item+ "%' ";          
            //DataTable dt = DbAccess.SelectBySql(sqlcheckitem).Tables[0];
            //if (dt != null && !string.IsNullOrEmpty(dt.Rows[0]["item"].ToString()))
            //{
            //    string ss = dt.Rows[0]["item"].ToString();
            //    int index = ss.IndexOf('F');
            //    if (index == ss.Length - 1)
            //    {
            //        item += "1";
            //    }
            //    else
            //    {
            //        string a = ss.Substring(index + 1);
            //        item += (int.Parse(a) + 1).ToString();
            //    }
            //}
            //else
            //{
            //    item += "1";
            //}

            string sqlcheckitem = @"  select max(convert(int,SUBSTRING(item,13,len(item) - 12 ))) item  from IPQCExceptionList where item like '%" + item + "%'  ";
            DataTable dt = DbAccess.SelectBySql(sqlcheckitem).Tables[0];
            if (dt != null && !string.IsNullOrEmpty(dt.Rows[0]["item"].ToString()))
            {         
                 string a = dt.Rows[0]["item"].ToString();
                 item += (int.Parse(a) + 1).ToString();
                
            }
            else
            {
                item += "1";
            }

            string ifrepeat = "否";
            string ifClose =  "否";
            string ifOvertime = "";

            string sqlifrepeat = "   select 1 from IPQCExceptionList where   org_id = '"+txtorg_id.Text+"'  and workno = '"+txtworkno.Text+"' and model = '"+ txtmodel.Text+ "' and  badclass = '"+txtbadclass.Text+ "'  ";
            DataTable repeatdt = DbAccess.SelectBySql(sqlifrepeat).Tables[0];
            if (repeatdt != null && repeatdt.Rows.Count >0)
            {
                ifrepeat = "是";
            }
            else
            {
                ifrepeat = "否";
            }

            if (txtbaddescribe.Text.Trim() != "" && txttemporaryhandle.Text.Trim() != "" && txtCauseAnalysis.Text.Trim() != "" && txtimprovemeasures.Text.Trim() != "")
            {
                ifClose = "是";
                ifOvertime = "否";
            }

            if (txtbadpicture.Text != "")
            {
              string floerPath = "";
              floerPath = "\\\\" + this.serverFilePath + "\\制程不良图片" + "\\" + txtworkno.Text + "\\" + item;
              if (floerPath != "")
                del_prefile(floerPath, "");
               bool falg = UploadReportFile(txtbadpicture.Text, "制程不良图片", txtworkno.Text, item);

               if (falg == false)
               {
                return;
               }
            }

            string msg = IPQCExceptioninsertList("保存", item, txtorg_id.Text, txtworkno.Text.Trim(),int.Parse(txtworknoqty.Text == "" ? "0" : txtworknoqty.Text), txthytcode.Text.Trim(),txtmodel.Text,int.Parse(txtqty.Text),int.Parse(txtNGQty.Text),ifrepeat,
                         txtstandid.Text, txtproductlineid.Text, txtcustomer.Text,txtbadclass.Text, txtbaddescribe.Text,txtproblemtype.Text, txttemporaryhandle.Text,
                         txtCauseAnalysis.Text, txtimprovemeasures.Text, txtdutyDepartment.Text,"",Login.username,txtchargeMan.Text,ifClose,ifOvertime,txtoverdutyDepartment.Text); //(byte[])object

          if (msg.Contains("保存成功"))
            {
                MessageBox.Show("保存成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                saveselect(item);
                resetcontrol();
            }
            else
            {
                MessageBox.Show("保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void sBtnupdate_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();
            string item = gridView.GetFocusedRowCellValue("表单编号").ToString();
            string hytcode = gridView.GetFocusedRowCellValue("Hytera编码").ToString();
            string ifrepeat = gridView.GetFocusedRowCellValue("是否重复发生").ToString();
            string ifOvertime = gridView.GetFocusedRowCellValue("是否超期").ToString();
            string ifClose = "否";
           
            if (txtbaddescribe.Text.Trim() !="" && txttemporaryhandle.Text.Trim() != "" && txtCauseAnalysis.Text.Trim() != "" && txtimprovemeasures.Text.Trim() != "")
            {
                ifClose = "是";
                ifOvertime = "否";
            }
            
      
            if (txtbadpicture.Text != "")
           {                        
            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\制程不良图片" + "\\" + txtworkno.Text + "\\" + item;
            if (floerPath != "")
                del_prefile(floerPath, "");
            bool falg = UploadReportFile(txtbadpicture.Text, "制程不良图片", txtworkno.Text, item);

            if (falg == false)
            {
                return;
            }
          }
            string msg = IPQCExceptioninsertList("保存", item,txtorg_id.Text, workno, int.Parse(txtworknoqty.Text == "" ? "0" : txtworknoqty.Text), hytcode, txtmodel.Text, int.Parse(txtqty.Text), int.Parse(txtNGQty.Text), ifrepeat,
                         txtstandid.Text, txtproductlineid.Text, txtcustomer.Text, txtbadclass.Text, txtbaddescribe.Text, txtproblemtype.Text, txttemporaryhandle.Text,
                         txtCauseAnalysis.Text, txtimprovemeasures.Text, txtdutyDepartment.Text, "",Login.username, txtchargeMan.Text, ifClose, ifOvertime, txtoverdutyDepartment.Text); //(byte[])object
            if (msg.Contains("保存成功"))
            {
                MessageBox.Show("更新成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                saveselect(item);
                resetcontrol();
            }
            else
            {
                MessageBox.Show("更新失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择要删除的行记录","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();
            string item = gridView.GetFocusedRowCellValue("表单编号").ToString();
            string hytcode = gridView.GetFocusedRowCellValue("Hytera编码").ToString();

            string sql = @"  delete IPQCExceptionList where item ='"+item+ "' and  workno = '"+workno+ "' and  hytcode= '"+ hytcode + "'  ";
            bool flag = DbAccess.ExecuteSql(sql);
            if (flag == true)
            {
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gridControl.DataSource = null;
            }
            else
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        void resetcontrol()
        {
            txtworkno.Text = "";
            txtworkno.ReadOnly = false;
            txthytcode.Text = "";
            txtmodel.Text = "";
            txtmodel.ReadOnly = false;
            txtworknoqty.Text = "";
            txtworknoqty.ReadOnly = true;
            txtcustomer.Text = "";
            txtcustomer.ReadOnly = false;
            txtstandid.Text = "";
            txtproductlineid.Text = "";
            txtbadclass.Text = "";
            txtdutyDepartment.Text = "";
            txtchargeMan.Text = "";
            txtoverdutyDepartment.Text = "";
            txtqty.Text = "";
            txtNGQty.Text = "";
            txtproblemtype.Text = "";
            txtbaddescribe.Text = "";
            txttemporaryhandle.Text = "";
            txtCauseAnalysis.Text = "";
            txtimprovemeasures.Text = "";
            txtbadpicture.Text = "";
            picbadImage.Image = null;
            txtIPQCExcepreport.Text = "";
        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            resetcontrol();
            gridControl.DataSource = null;

        }

        void selectdata()
        {
            string where = " where 1=1 ";
            string workno = txtworkno.Text.Trim();
            string hytcode = txthytcode.Text;
            string customer = txtcustomer.Text;
            if (!string.IsNullOrEmpty(workno))
            {
                where += " and workno = '" + workno + "' ";
            }
            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer = '" + customer + "' ";
            }
            if (!string.IsNullOrEmpty(hytcode))
            {
                where += " and hytcode = '" + hytcode + "' ";
            }
            string sql = @"  select  ifrepeat 是否重复发生,standid 站别,item 表单编号,checkdate 日期,productlineid 生产线别,customer 客户,org_id 组织,
		        workno 工单号,worknoqty 工单数量,hytcode Hytera编码,model 客户机型,qty 投入数,NGQty 不良数,CASE WHEN NGQty < worknoqty  then  convert( varchar,convert(numeric(3,1),(NGQty+0.0)/(isnull(worknoqty,qty))*100 ))+'%' ELSE '100%' end 不良率,
				badclass 不良类别,baddescribe 问题描述,problemtype 问题分类,temporaryhandle 临时处理方法,CauseAnalysis 原因分析,improvemeasures 改善计划,
				dutyDepartment 责任部门,updateMan 记录人,chargeMan 责任人,ifClose 是否关闭,ifOvertime 是否超期,overdutyDepartment 延时责任部门,badImage 不良图片
		        from IPQCExceptionList  ";
            sql += where;
            sql += "  order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
                repositoryItemMemoEdit1.LinesCount = 3;
            }

        }
        private void sBtnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string workno = txtworkno.Text.Trim();
            string standid = txtstandid.Text;
            string productlineid = txtproductlineid.Text;
            string badclass = txtbadclass.Text;
            string model = txtmodel.Text;
            string customer = txtcustomer.Text;


            if (!string.IsNullOrEmpty(workno))
            {
                where += " and workno = '" + workno + "' ";
            }
            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer = '" + customer + "' ";
            }
            if (!string.IsNullOrEmpty(model))
            {
                where += " and model like '%" + model + "%' ";
            }
            if (!string.IsNullOrEmpty(standid))
            {
                where += " and standid = '" + standid + "' ";
            }
            if (!string.IsNullOrEmpty(productlineid))
            {
                where += " and productlineid = '" + productlineid + "' ";
            }
            if (!string.IsNullOrEmpty(badclass))
            {
                where += " and badclass = '" + badclass + "' ";
            }

            string sql = @"  select  ifrepeat 是否重复发生,standid 站别,item 表单编号,checkdate 日期,productlineid 生产线别,customer 客户,org_id 组织,
		        workno 工单号,worknoqty 工单数量,hytcode Hytera编码,model 客户机型,qty 投入数,NGQty 不良数,CASE WHEN NGQty < worknoqty  then  convert( varchar,convert(numeric(3,1),(NGQty+0.0)/(isnull(worknoqty,qty))*100 ))+'%' ELSE '100%' end 不良率,
				badclass 不良类别,baddescribe 问题描述,problemtype 问题分类,temporaryhandle 临时处理方法,CauseAnalysis 原因分析,improvemeasures 改善计划,
				dutyDepartment 责任部门,updateMan 记录人,chargeMan 责任人,ifClose 是否关闭,ifOvertime 是否超期,overdutyDepartment 延时责任部门,badImage 不良图片
		        from IPQCExceptionList  ";
            sql += where;

            sql += "  order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
                repositoryItemMemoEdit1.LinesCount = 3;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }
        }

        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "制程异常信息记录";
            int n = name.LastIndexOf(".") + 1;
            if (n > 0) name = name.Substring(n, name.Length - n);
            dlg.Title = "导出 " + title;
            dlg.FileName = name;
            dlg.Filter = filter;
            if (dlg.ShowDialog() == DialogResult.OK) return dlg.FileName;
            return "";
        }
        private void ExportToEx(String filename, string ext, BaseView exportView)
        {
            Cursor currentCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            if (ext == "rtf") exportView.ExportToRtf(filename);
            if (ext == "pdf") exportView.ExportToPdf(filename);
            if (ext == "mht") exportView.ExportToMht(filename);
            if (ext == "htm") exportView.ExportToHtml(filename);
            if (ext == "txt") exportView.ExportToText(filename);
            if (ext == "xls") exportView.ExportToXls(filename);
            if (ext == "xlsx") exportView.ExportToXlsx(filename);
            Cursor.Current = currentCursor;
        }
        private void OpenFile(string fileName)
        {
            if (XtraMessageBox.Show("Do you want to open this file?", "Export To...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.StartInfo.FileName = fileName;
                    process.StartInfo.Verb = "Open";
                    process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                    process.Start();
                }
                catch
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(this, "Cannot find an application on your system suitable for openning the file with exported data.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            progressBarControl1.Position = 0;
        }
        void gridView_PrintExportProgress(object sender, ProgressChangedEventArgs e)
        {
            SetPosition(e.ProgressPercentage);
        }
        void SetPosition(int pos)
        {
            progressBarControl1.Position = pos;
            this.Update();
        }

        private void sBtnreport_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }

        private void browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "文件(*.pdf)|*.pdf";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;
            this.txtIPQCExcepreport.Text = "";
            this.txtIPQCExcepreport.Text = ofdImport.FileName;
        }

        public bool Connect(string remoteHost)
        {
            bool Flag = true;
            Process proc = new Process();
            try
            {
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                string dosdel = @"net use \\" + remoteHost + " /del";
                proc.StandardInput.WriteLine(dosdel);
                //string dosLine = @"net use \\" + remoteHost + " hytera;2012" + " /user:" + "Upload";
                string dosLine = @"net use \\" + remoteHost + " hytera;2012" + " /user:" + this.serverFilePath.Split('\\')[0] + "\\Upload";
                proc.StandardInput.WriteLine(dosLine);
                proc.StandardInput.WriteLine("exit");
                while (proc.HasExited == false)
                {
                    proc.WaitForExit(1000);
                }
                string errormsg = proc.StandardError.ReadToEnd();
                if (errormsg != "")
                {
                    Flag = false;
                }
                proc.StandardError.Close();
            }
            catch (Exception ex)
            {
                Flag = false;
            }
            finally
            {
                try
                {
                    proc.Close();
                    proc.Dispose();
                }
                catch
                {
                }
            }
            return Flag;
        }


        protected void del_prefile(string filepath, string filename)
        {
            if (Directory.Exists(filepath))
            {
                if (Connect(serverFilePath))
                {
                    string[] Mulfile = Directory.GetFiles(filepath);
                    foreach (string ss in Mulfile)
                    {
                        if (System.IO.File.Exists(ss))
                            System.IO.File.Delete(ss);
                    }
                }
            }

        }

        private bool UploadReportFile(string filepath, string directype, string workno, string item)
        {
            bool b = false;
            if (Connect(serverFilePath))
            {
                bool re = true;
                if (filepath != "")
                {
                    try
                    {
                        string[] fileName = @filepath.Split('\\');
                        string floerPath = "\\\\" + this.serverFilePath + "\\" + directype + "\\" + workno + "\\" + item;
                        string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];
                        if (Directory.Exists(floerPath) == false)   //如果不存在就创建file文件夹
                        {
                            Directory.CreateDirectory(floerPath);
                        }
                        File.Copy(filepath, fileServerPath, true);
                        re = true;
                    }
                    catch (Exception e)
                    {
                        re = false;
                    }
                }
                if (re)
                {
                   ////// MessageBox.Show("文件上传成功");
                    b = true;
                }
                else MessageBox.Show(filepath + " 文件上传失败");
            }
            else
            {
                MessageBox.Show("无法连接到服务器的共享目录");
            }
            return b;
        }

        private void btnupload_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (txtIPQCExcepreport.Text.Trim() == "")
            {
                return;
            }
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择相应的行记录", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();
            string item = gridView.GetFocusedRowCellValue("表单编号").ToString();
            string hytcode = gridView.GetFocusedRowCellValue("Hytera编码").ToString();

            string fileDBServerPath = "";
            if (txtIPQCExcepreport.Text != "")
            {
                string s = txtIPQCExcepreport.Text;
                string[] fileName = @s.Split('\\');
                fileDBServerPath += "\\\\" + this.serverFilePath + "\\制程异常处理报告" + "\\" + workno + "\\" + item + "\\" + fileName[fileName.Length - 1];
            }
            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\制程异常处理报告" + "\\" + workno + "\\" + item;
            if (floerPath != "")
                del_prefile(floerPath, "");
            bool b = UploadReportFile(txtIPQCExcepreport.Text, "制程异常处理报告", workno, item);
            if (b)
            {
                string sqlupload = "   update IPQCExceptionList set ifClose = '是' ,ifovertime = case when ifovertime ='' and datediff(hour,checkdate,GETDATE())>48 then '是' when ifovertime ='' and datediff(hour,checkdate,GETDATE())<=48 then '否' ";
                       sqlupload  +="  else ifovertime end  where item = '"+item+ "' and workno = '"+workno+ "'  ";
                bool falg = DbAccess.ExecuteSql(sqlupload);
                if (falg)
                {
                    MessageBox.Show("上传成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }               
            }
        }

        private void OpenFile(string filepath, string pdffile)
        {
            string filename = "";
            //filename = pdffile + ".pdf";
            filename = pdffile;
            //定义一个ProcessStartInfo实例
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            //设置启动进程的初始目录
            //string[] fileName = pdffile;

            //filename = fileName[fileName.Length - 1];
            info.FileName = @filename;
            info.WorkingDirectory = filepath;
            //设置启动进程的参数
            info.Arguments = "";
            //启动由包含进程启动信息的进程资源
            try
            {
                System.Diagnostics.Process.Start(info);
                //System.Diagnostics.Process.Start(Application.StartupPath +"\\"+info);

            }

            catch (System.ComponentModel.Win32Exception we)
            {

                MessageBox.Show(this, we.Message);
                return;
            }

        }
        void export(string item)
        {
            string sql = "";

            int sheetCount = 1;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\制程异常处理单表格.xlsx";
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            if (sheet == null)
                return;

            sql = @"  select  ifrepeat 是否重复发生,standid 站别,item 表单编号,checkdate 日期,productlineid 生产线别,customer 客户,
		        workno 工单号,hytcode Hytera编码,model 客户机型,qty 投入数,NGQty 不良数,CASE WHEN NGQty < worknoqty  then  convert( varchar,convert(numeric(3,1),(NGQty+0.0)/(isnull(worknoqty,qty))*100 ))+'%' ELSE '100%' end 不良率,
				badclass 不良类别,baddescribe 问题描述,problemtype 问题分类,temporaryhandle 临时处理方法,CauseAnalysis 原因分析,improvemeasures 改善计划,
				dutyDepartment 责任部门,chargeMan 责任人,ifClose 是否关闭,ifOvertime 是否超期,overdutyDepartment 延时责任部门,badImage 不良图片
		        from IPQCExceptionList  where item = '" + item + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null || dt.Rows.Count < 1)
                return;

            string ifrepeat = dt.Rows[0]["是否重复发生"].ToString();
            string problemtype = dt.Rows[0]["问题分类"].ToString();
            if (ifrepeat == "是")
            {
                sheet.Cells.get_Range("D22").Value = "√是";
            }
            else
            {
                sheet.Cells.get_Range("E22").Value = "√否";
            }

            /*
            if (problemtype == "物料")
            {
                sheet.Cells.get_Range("C21").Value = "√";
            }
            else if (problemtype == "作业")
            {
                sheet.Cells.get_Range("E21").Value = "√";
            }
            else if (problemtype == "工艺")
            {
                sheet.Cells.get_Range("G21").Value = "√";
            }
            else if (problemtype == "设计")
            {
                sheet.Cells.get_Range("I21").Value = "√";
            }
            else if (problemtype == "设备")
            {
                sheet.Cells.get_Range("K21").Value = "√";
            }
            */

            sheet.Cells.get_Range("H2").Value = dt.Rows[0]["表单编号"].ToString();
            sheet.Cells.get_Range("C3").Value = dt.Rows[0]["客户机型"].ToString();
            sheet.Cells.get_Range("F3").Value = dt.Rows[0]["Hytera编码"].ToString();
            sheet.Cells.get_Range("J3").Value = dt.Rows[0]["生产线别"].ToString();
            sheet.Cells.get_Range("C4").Value = dt.Rows[0]["投入数"].ToString();
            sheet.Cells.get_Range("F4").Value = dt.Rows[0]["不良数"].ToString();
            sheet.Cells.get_Range("J4").Value = dt.Rows[0]["不良率"].ToString();
            sheet.Cells.get_Range("A5").Value = "异常描述:" + "\r\n" + dt.Rows[0]["问题描述"].ToString();
            sheet.Cells.get_Range("D10").Value = dt.Rows[0]["临时处理方法"].ToString();
            sheet.Cells.get_Range("B15").Value = dt.Rows[0]["原因分析"].ToString();
            sheet.Cells.get_Range("B21").Value = "问题分类：" + dt.Rows[0]["问题分类"].ToString();
            sheet.Cells.get_Range("B24").Value = dt.Rows[0]["改善计划"].ToString();


        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)             
              return;    
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();
            string item = gridView.GetFocusedRowCellValue("表单编号").ToString();
            string hytcode = gridView.GetFocusedRowCellValue("Hytera编码").ToString();

            if (e.Column.FieldName == "表单编号")
            {

                string floerPath = "\\\\" + serverFilePath + "\\制程异常处理报告" + "\\" + workno + "\\" + item;

                if (Connect(serverFilePath))
                {
                    if (Directory.Exists(floerPath) == false)
                    {
                        MessageBox.Show("还没有上传相应的处理报告!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    string[] str = Directory.GetFiles(floerPath);
                    if (str.Length > 0)
                    {
                        for (int i = 0; i < str.Length; i++)
                        {
                            if (str[i].ToString().Contains(".db"))
                                continue;
                            string filename = Path.GetFileName(str[i].ToString());

                            FileStream fs = new FileStream(str[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            BinaryReader reader = new BinaryReader(fs);
                            string fileclass = "";
                            try
                            {
                                for (int j = 0; j < 2; j++)
                                {
                                    fileclass += reader.ReadByte().ToString();
                                }
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            {
                                OpenFile(floerPath, filename);
                                fs.Close();
                                fs.Dispose();
                                reader.Close();
                                return;
                            }
                        }

                    }
                }
            }
            if (e.Column.FieldName == "工单号")
            {
                try
                {
                    export(item);

                }
                catch (Exception ex)
                {

                    MessageBox.Show("还没生成完整报表", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            if (e.Column.FieldName == "Hytera编码")
            {
                try
                {                   
                    string badfilePath = "\\\\" + serverFilePath + "\\制程不良图片" + "\\" + workno + "\\" + item;
                    Cursor.Current = Cursors.WaitCursor;
                    if (Connect(serverFilePath))
                    {                     
                        if (Directory.Exists(badfilePath))
                        {
                            string[] pt = Directory.GetFiles(badfilePath);
                            if (pt.Length == 0)
                                return;
                            if (pt.Length > 0)
                            {
                                for (int i = 0; i < pt.Length; i++)
                                {
                                    if (pt[i].ToString().Contains(".db"))
                                        continue;
                                    FileStream ft = new FileStream(pt[i].ToString(), FileMode.Open);
                                    Image badpic = Image.FromStream(ft);
                                    ft.Close();
                                    ft.Dispose();
                                    picbadImage.Image = badpic;
                                    picbadImage.Properties.SizeMode = PictureSizeMode.Stretch;
                                }

                            }
                        }
                    }
                    Cursor.Current = Cursors.Default;

                }
                catch (Exception ex)
                {

                    MessageBox.Show("还没上传不良图片", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

            try
            {
                txtworkno.Text = gridView.GetFocusedRowCellValue("工单号").ToString();
                txtworknoqty.Text = gridView.GetFocusedRowCellValue("工单数量").ToString();
                txthytcode.Text = gridView.GetFocusedRowCellValue("Hytera编码").ToString();
                txtmodel.Text = gridView.GetFocusedRowCellValue("客户机型").ToString();
                txtcustomer.Text = gridView.GetFocusedRowCellValue("客户").ToString();
                txtstandid.Text = gridView.GetFocusedRowCellValue("站别").ToString();
                txtproductlineid.Text = gridView.GetFocusedRowCellValue("生产线别").ToString();
                txtbadclass.Text = gridView.GetFocusedRowCellValue("不良类别").ToString();
                txtdutyDepartment.Text = gridView.GetFocusedRowCellValue("责任部门").ToString();
                txtchargeMan.Text = gridView.GetFocusedRowCellValue("责任人").ToString();
                txtoverdutyDepartment.Text = gridView.GetFocusedRowCellValue("延时责任部门").ToString();                             
                txtqty.Text = gridView.GetFocusedRowCellValue("投入数").ToString();
                txtNGQty.Text = gridView.GetFocusedRowCellValue("不良数").ToString();
                txtproblemtype.Text = gridView.GetFocusedRowCellValue("问题分类").ToString();
                txtbaddescribe.Text = gridView.GetFocusedRowCellValue("问题描述").ToString();
                txttemporaryhandle.Text = gridView.GetFocusedRowCellValue("临时处理方法").ToString();
                txtCauseAnalysis.Text = gridView.GetFocusedRowCellValue("原因分析").ToString();
                txtimprovemeasures.Text = gridView.GetFocusedRowCellValue("改善计划").ToString();

                sBtnsave.Enabled = false;

            }
            catch
            {

            }
        }

        private void btnuploadbadpic_Click(object sender, EventArgs e)
        {

        }


        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            //DataTable dt = gridControl.DataSource as DataTable;
            //if (dt == null || dt.Rows.Count < 1)
            //{
            //    return;
            //}

            //if (float.Parse( gridView.GetDataRow(e.RowHandle)["不良率"].ToString().Replace("%", "")) >= 10)
            //{
            //    e.Appearance.BackColor = Color.Red;
            //}
        }
        public DateTime beforeTime, afterTime;

        private void IPQCExceptionList_FormClosing(object sender, FormClosingEventArgs e)
        {
            Process[] myProcesses;
            DateTime startTime;
            myProcesses = Process.GetProcessesByName("Excel");

            // 得不到Excel进程ID，暂时只能判断进程启动时间 
            foreach (Process myProcess in myProcesses)
            {
                startTime = myProcess.StartTime;

                if (startTime > beforeTime && startTime < afterTime)
                {
                    myProcess.Kill();
                }
            }
        }
    }
}