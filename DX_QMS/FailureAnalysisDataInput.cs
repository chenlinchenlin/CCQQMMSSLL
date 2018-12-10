using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class FailureAnalysisDataInput : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public FailureAnalysisDataInput()
        {
            InitializeComponent();
            SetRule();
        }

        private void SetRule()
        {
            string post = "";
            if (Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "失效分析");
            btn_Search.Enabled = dic["hasQuery"];
            btn_Save.Enabled = dic["hasInsert"];
            btn_Delete.Enabled = dic["hasDelete"];

        }

        private void FailureAnalysisDataInput_Load(object sender, EventArgs e)
        {
            // 紧急程度 初始化
            cbox_Emergency.Properties.Items.Add("急单");
            cbox_Emergency.Properties.Items.Add("普通");
            cbox_Emergency.SelectedIndex = -1;

            // 完成状态 初始化
            cbox_Status.Properties.Items.Add("已完成");
            cbox_Status.Properties.Items.Add("待完成");
            cbox_Status.SelectedIndex = -1;
        }

        public DataSet FailureAnalysisDataInputInfo(string operate, string ProjectName, string SN, string ApplyDate, string Customer, string Emergency, string FinishDate,
    int ProjectQty, string Status)
        {
            SqlParameter[] para = new SqlParameter[10];
            para[0] = new SqlParameter("@operate", operate);
            para[1] = new SqlParameter("@ProjectName", ProjectName);
            para[2] = new SqlParameter("@SN", SN);
            para[3] = new SqlParameter("@ApplyDate", ApplyDate);
            para[4] = new SqlParameter("@Customer", Customer);
            para[5] = new SqlParameter("@Emergency", Emergency);
            para[6] = new SqlParameter("@FinishDate", FinishDate);
            para[7] = new SqlParameter("@ProjectQty", ProjectQty);
            para[8] = new SqlParameter("@Status", Status);
            para[9] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[9].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "FailureAnalysisDataInput", para);
        }
        private void LimitClose()
        {
            txt_Customer.ReadOnly = true;
            txt_ProjectQty.ReadOnly = true;
            txt_Projects.ReadOnly = true;
            txt_SN.ReadOnly = true;
            cbox_Emergency.Enabled = false;
            dtpick_Apply.Enabled = false;
            dtpick_Finish.Enabled = false;
        }

        private void txt_SN_Leave(object sender, EventArgs e)
        {
            if (txt_SN.Text != "")
            {
                DataSet ds = FailureAnalysisDataInputInfo("获取编号数据", "", txt_SN.Text, "", "", "", "", 0, "");
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    // 控件赋值数据
                    txt_Projects.Text = ds.Tables[0].Rows[0]["ExperimentProject"].ToString();
                    dtpick_Apply.Text = ds.Tables[0].Rows[0]["ApplicationDate"].ToString();
                    txt_Customer.Text = ds.Tables[0].Rows[0]["Customer"].ToString();
                    cbox_Emergency.Text = ds.Tables[0].Rows[0]["LevelState"].ToString();
                    dtpick_Finish.Text = ds.Tables[0].Rows[0]["RequiredCompleteDate"].ToString();
                    txt_ProjectQty.Text = ds.Tables[0].Rows[0]["ProjectQty"].ToString();
                    dtpick_Apply.Checked = true;
                    dtpick_Finish.Checked = true;

                    // 关闭输入功能
                    LimitClose();
                }
            }
        }

        private void txt_ProjectQty_Leave(object sender, EventArgs e)
        {
            int temp = 0;
            if (int.TryParse(txt_ProjectQty.Text, out temp) == false)
            {
                MessageBox.Show("实验项目数必须为数字类型", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txt_ProjectQty.Text = "";
                return;
            }
            if (Convert.ToInt32(txt_ProjectQty.Text) < 0)
            {
                MessageBox.Show("实验项目数不能小于0", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txt_ProjectQty.Text = "";
                return;
            }
        }

        public string FailureAnalysisDatainput(string operate, string ProjectName, string SN, string ApplyDate, string Customer, string Emergency, string FinishDate,
            int ProjectQty, string Status)
        {
            SqlParameter[] para = new SqlParameter[10];
            para[0] = new SqlParameter("@operate", operate);
            para[1] = new SqlParameter("@ProjectName", ProjectName);
            para[2] = new SqlParameter("@SN", SN);
            para[3] = new SqlParameter("@ApplyDate", ApplyDate);
            para[4] = new SqlParameter("@Customer", Customer);
            para[5] = new SqlParameter("@Emergency", Emergency);
            para[6] = new SqlParameter("@FinishDate", FinishDate);
            para[7] = new SqlParameter("@ProjectQty", ProjectQty);
            para[8] = new SqlParameter("@Status", Status);
            para[9] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[9].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "FailureAnalysisDataInput", para);
            return para[9].Value.ToString();
        }

        void selectdata(string SNO)
        {
            string sql = @"select * from SMT_ExperimentProject where FileNo ='"+ SNO + "'";
            DataTable  ds = DbAccess.SelectBySql(sql).Tables[0];
            if (ds != null && ds.Rows.Count > 0)
            {
                dataGridView1.DataSource = ds;
            }

        }

        private void LimitOpen()
        {
            txt_Customer.ReadOnly = false;
            txt_ProjectQty.ReadOnly = false;
            txt_Projects.ReadOnly = false;
            txt_SN.ReadOnly = false;
            cbox_Emergency.Enabled = true;
            dtpick_Apply.Enabled = true;
            dtpick_Finish.Enabled = true;
        }

        private void ClearControl()
        {
            txt_Projects.Text = "";
            txt_Customer.Text = "";
            txt_ProjectQty.Text = "";
            txt_SN.Text = "";
            cbox_Emergency.SelectedIndex = -1;
            cbox_Status.SelectedIndex = -1;
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            if (txt_Customer.Text != "" && txt_ProjectQty.Text != "" && txt_Projects.Text != "" && txt_SN.Text != "" && cbox_Status.Text != "" && cbox_Emergency.Text != ""
    && dtpick_Apply.Checked == true && dtpick_Finish.Checked == true)
            {
                string msg = FailureAnalysisDatainput("保存", txt_Projects.Text, txt_SN.Text, dtpick_Apply.Text, txt_Customer.Text, cbox_Emergency.Text,
                    dtpick_Finish.Text, Convert.ToInt32(txt_ProjectQty.Text), cbox_Status.Text);
                if (msg.Contains("成功"))
                {
                    MessageBox.Show(msg, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (!msg.Contains("更新"))
                    {

                        //// 新增一列
                        //int count = dataGridView1.Rows.Count;
                        //dataGridView1.Rows.Add(1);
                        //dataGridView1.Rows[count].Cells["实验项目"].Value = txt_Projects.Text;
                        //dataGridView1.Rows[count].Cells["申请时间"].Value = dtpick_Apply.Text;
                        //dataGridView1.Rows[count].Cells["编号"].Value = txt_SN.Text;
                        //dataGridView1.Rows[count].Cells["客户"].Value = txt_Customer.Text;
                        //dataGridView1.Rows[count].Cells["紧急程度"].Value = cbox_Emergency.Text;
                        //dataGridView1.Rows[count].Cells["要求完成时间"].Value = dtpick_Finish.Text;
                        //dataGridView1.Rows[count].Cells["实验项目数"].Value = txt_ProjectQty.Text;
                        //dataGridView1.Rows[count].Cells["完成状态"].Value = cbox_Status.Text;
                        string SNO = txt_SN.Text;
                        selectdata(SNO);

                    }
                    else
                    {
                        //// 循环查找，更新
                        //for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        //{
                        //    if (dataGridView1.Rows[i].Cells["编号"].Value.ToString() == txt_SN.Text)
                        //    {
                        //        dataGridView1.Rows[i].Cells["完成状态"].Value = cbox_Status.Text;
                        //        break;
                        //    }
                        //}
                        string SNO = txt_SN.Text;
                        selectdata(SNO);
                    }

                    // 清空控件
                    LimitOpen();
                    ClearControl();
                }
            }
            else
            {
                MessageBox.Show("数据不齐，请补充完整!", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        private void btn_Search_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            string sqlstr = "select * from SMT_ExperimentProject where 1 = 1 ";
            if (txt_Projects.Text != "")
            {
                sqlstr += "and ExperimentProject = '" + txt_Projects.Text + "' ";
            }
            if (txt_SN.Text != "")
            {
                sqlstr += " and FileNo = '" + txt_SN.Text + "' ";
            }
            if (dtpick_Apply.Checked == true)
            {
                sqlstr += " and ApplicationDate = '" + dtpick_Apply.Text + "' ";
            }
            if (txt_Customer.Text != "")
            {
                sqlstr += " and Customer = '" + txt_Customer.Text + "' ";
            }
            if (cbox_Emergency.Text != "")
            {
                sqlstr += " and LevelState = '" + cbox_Emergency.Text + "' ";
            }
            if (dtpick_Finish.Checked == true)
            {
                sqlstr += " and RequiredCompleteDate = '" + dtpick_Finish.Text + "' ";
            }
            if (txt_ProjectQty.Text != "")
            {
                sqlstr += " and ProjectQty = " + Convert.ToInt32(txt_ProjectQty.Text) + " ";
            }
            if (cbox_Status.Text != "")
            {
                sqlstr += " and CompleteState = '" + cbox_Status.Text + "' ";
            }

            DataSet ds = DbAccess.SelectBySql(sqlstr);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                dataGridView1.DataSource = ds.Tables[0];
            }
            else
            {
                MessageBox.Show("没有数据！","提醒",MessageBoxButtons.OK,MessageBoxIcon.Warning);

            }
        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            LimitOpen();
            ClearControl();
            dataGridView1.DataSource = null;

        }
        private void btn_Delete_Click(object sender, EventArgs e)
        {
            DataTable dt = dataGridView1.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string msg = FailureAnalysisDatainput("删除", "", gridView.GetFocusedRowCellValue("FileNo").ToString(), "", "", "", "", 0, "");
            if (msg.Contains("成功"))
            {
                MessageBox.Show(gridView.GetFocusedRowCellValue("FileNo").ToString() + "删除成功", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gridView.DeleteRow(gridView.FocusedRowHandle);
                return;
            }

        }

        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "失效分析数据维护";
            int n = name.LastIndexOf(".") + 1;
            if (n > 0) name = name.Substring(n, name.Length - n);
            dlg.Title = "Export To " + title;
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

        private void btn_Import_Click(object sender, EventArgs e)
        {
            DataTable dt = dataGridView1.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);

        }
    }
}