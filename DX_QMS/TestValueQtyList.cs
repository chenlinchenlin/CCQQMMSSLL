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

namespace DX_QMS
{
    public partial class TestValueQtyList : DevExpress.XtraEditors.XtraForm
    {
        Common.IQC ic = new Common.IQC();
        int locY1 = 2, counter01 = 0, locX1 = 10;
        DataTable dt = null;
        string lotnotemp = "";
        public int testqty = 0, lotqtytemp = 0;
        TextBox t;
        public DataSet ds = null;
        private string ssamplefacetqy = "0";
        public string sfq
        {
            get { return ssamplefacetqy; }
            set { this.ssamplefacetqy = value; }
        }
        public TestValueQtyList()
        {
            InitializeComponent();
        }

        private void TestValueQtyList_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            this.Close();
        }

        private void btnclear_Click(object sender, EventArgs e)
        {
            clearTextBox();
        }

        private void btnsave_Click(object sender, EventArgs e)
        {

            bool ifSucess = false;
            for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
            {
                string[] mssg = new string[ds.Tables[0].Rows.Count];

                string msg = ""; int i = 1;
                double scope1 = 0, scope2 = 0;
                int k = 0; string m = "";

                foreach (Control tb in this.panel1.Controls)
                {
                    if (tb is Label) continue;
                    if (tb is TextBox && tb.Name.Contains(ds.Tables[0].Rows[j]["TestSubItem"].ToString()))
                    {
                        bool s1 = false, s2 = false;
                        if (tb.Text == "")
                        {
                            MessageBox.Show("值不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            tb.Focus();
                            return;
                        }
                        double n = 0;
                        if (double.TryParse(tb.Text, out n) == false)
                        {
                            MessageBox.Show(tb.Text + "值错误,请用数字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            tb.Focus();
                            tb.Text = "";
                            return;
                        }
                        foreach (Control s in this.Controls)
                        {
                            if (s.Text.Contains(tb.Name.Substring(0, tb.Name.IndexOf(":"))) && s.Text.Contains("范围"))
                            {
                                if (s is Label && s.Text.Contains("范围1"))
                                {
                                    scope1 = double.Parse(s.Text.Substring(s.Text.IndexOf("范围1:") + 4) == "" ? "0" : s.Text.Substring(s.Text.IndexOf("范围1:") + 4));
                                    s1 = true;
                                    if (s2 == true)
                                        break;
                                    else
                                        continue;
                                }
                                else if (s is Label && s.Text.Contains("范围2"))
                                {
                                    scope2 = double.Parse(s.Text.Substring(s.Text.IndexOf("范围2:") + 4) == "" ? "0" : s.Text.Substring(s.Text.IndexOf("范围2:") + 4));
                                    s2 = true;
                                    if (s1 == true)
                                        break;
                                    else
                                        continue;
                                }
                            }
                            else
                                continue;
                        }

                        try
                        {

                            m = ds.Tables[0].Rows[j]["TestSubItem"].ToString();
                            foreach (Control t in this.panel1.Controls)
                            {
                                if (t is TextBox)
                                {
                                    if (t.Name.Contains(m))
                                        k += 1;
                                }
                            }

                            if (double.Parse(tb.Text) < scope1 || double.Parse(tb.Text) > scope2)
                                //msg += "测试值" + i.ToString() + ":" + tb.Text + "[NG]" + ";";
                                mssg[j] += "测试值" + i.ToString() + ":" + tb.Text + "[NG]" + ";";
                            else
                                //msg += "测试值" + i.ToString() + ":" + tb.Text + "[OK]" + ";";
                                mssg[j] += "测试值" + i.ToString() + ":" + tb.Text + "[OK]" + ";";
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        i = i + 1;
                        if (i >= k)
                            break;
                        else
                            continue;
                    }
                }
                string msgbox = "";
                if (mssg[j].ToString().Contains("NG"))
                {
                    if (MessageBox.Show("规格:" + m + mssg[j], "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                    else
                    {

                        try
                        {
                            string sstate = "NG";

                            msgbox = ic.AddNewTestList("新增", ds.Tables[0].Rows[j]["testtype"].ToString(), ds.Tables[0].Rows[j]["TestItem"].ToString(), ds.Tables[0].Rows[j]["TestSubItem"].ToString(), ds.Tables[0].Rows[j]["TestDesc"].ToString(), ds.Tables[0].Rows[j]["TestTool"].ToString(), ds.Tables[0].Rows[j]["PackType"].ToString()
                               , ds.Tables[0].Rows[j]["SampleType"].ToString(), float.Parse(scope2.ToString()),
                              float.Parse(scope1.ToString()), float.Parse(ds.Tables[0].Rows[j]["AQLValue"].ToString() == "" ? "0" : ds.Tables[0].Rows[j]["AQLValue"].ToString()), lotnotemp, lotqtytemp,
                              Login.username, int.Parse(lblsampleqty.Text), 1, ds.Tables[0].Rows[j]["Productcode"].ToString(), ds.Tables[0].Rows[j]["AQL"].ToString(), sstate, "", "", 0, mssg[j].ToString(), lblunit.Text, ds.Tables[0].Rows[j]["CheckType"].ToString(), "", "", "", lblrsno.Text);
                            ifSucess = true;

                        }
                        catch (Exception ex)
                        {
                            throw new Exception(ex.Message);
                        }
                    }
                }
                else
                {
                    try
                    {
                        string sstate = "OK";

                        msgbox = ic.AddNewTestList("新增", ds.Tables[0].Rows[j]["testtype"].ToString(), ds.Tables[0].Rows[j]["TestItem"].ToString(), ds.Tables[0].Rows[j]["TestSubItem"].ToString(), ds.Tables[0].Rows[j]["TestDesc"].ToString(), ds.Tables[0].Rows[j]["TestTool"].ToString(), ds.Tables[0].Rows[j]["PackType"].ToString()
                                 , ds.Tables[0].Rows[j]["SampleType"].ToString(), float.Parse(scope2.ToString()),
                                float.Parse(scope1.ToString()), float.Parse(ds.Tables[0].Rows[j]["AQLValue"].ToString() == "" ? "0" : ds.Tables[0].Rows[j]["AQLValue"].ToString()), lotnotemp, lotqtytemp,
                                Login.username, int.Parse(lblsampleqty.Text), 1, ds.Tables[0].Rows[j]["Productcode"].ToString(), ds.Tables[0].Rows[j]["AQL"].ToString(), sstate, "", "", 0, mssg[j].ToString(), lblunit.Text, ds.Tables[0].Rows[j]["CheckType"].ToString(), "", "", "", lblrsno.Text);
                        ifSucess = true;

                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
            //表示此次操作是否有增加测试记录(ifSucess==True),因为如有多个子测试项,抽样数一次只能加1,否则的话,会加多次
            if (ifSucess)
            {
                MessageBox.Show("新增成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                clearTextBox();
                lblsamplefactqty.Text = (int.Parse(lblsamplefactqty.Text == "" ? "0" : lblsamplefactqty.Text) + 1).ToString();
                //把ifSucess复位,让其继续下一次测试,因为用户可能再次在该页面增加另一组测试记录
                ifSucess = false;
            }
            //if (lblsampleqty.Text == lblsamplefactqty.Text)
            {
                this.ssamplefacetqy = lblsamplefactqty.Text.Trim();
                this.Dispose();
                this.Close();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                if (ActiveControl is TextBox)
                {
                    System.Windows.Forms.SendKeys.Send("{tab}");
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void clearTextBox()
        {
            foreach (Control c in this.panel1.Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                }
            }
        }

        public TestValueQtyList(int testvalueqty, string scope1, string scope2, string sampleqty, string samplefactqty, string lotno, int lotqty, string testtype, string testitem, string productcode, DataTable dtnew, string RSNO)
        {
            InitializeComponent();

            dt = dtnew;
            testqty = testvalueqty;
            lblsampleqty.Text = sampleqty;
            lblsamplefactqty.Text = samplefactqty;
            lblunit.Text = dtnew.Rows[0]["unit"].ToString();
            lotnotemp = lotno;
            lotqtytemp = lotqty;
            this.lblrsno.Text = RSNO;
            string ssql = "select * from IQC_TestProgSet where TestItem='" + testitem + "' and TestType='" + testtype + "' and Productcode='" + productcode + "' order by SubItem";
            ds = Common.DbAccess.SelectBySql(ssql);

            for (int j = 1; j <= ds.Tables[0].Rows.Count; j++)
            {
                Label litem = new Label();

                ////litem.Font = new Font("宋体", 12);
                litem.Text = ds.Tables[0].Rows[j - 1]["TestSubItem"].ToString();

                //control.Font = new Font("华文新魏", 22.2f, FontStyle.Bold | Ityle.ItalicFontStyle.Underline);
               // litem.Font = new Font("宋体", 12);

                locY1 += litem.Height + 18;
                litem.Location = new Point(locX1, locY1);
                this.Controls.Add(litem);
                Label scop2 = new Label();

              ////  scop2.Font = new Font("宋体", 12);
                scop2.Text = litem.Text + "范围2: " + ds.Tables[0].Rows[j - 1]["UpValue"].ToString();
               // scop2.Font = new Font("宋体", 12);


                scop2.Name = litem.Text;
                locY1 -= 20;
                locX1 += 0;
                scop2.Location = new Point(locX1, locY1);
                this.Controls.Add(scop2);
                Label scop1 = new Label();

               ///// scop1.Font = new Font("宋体", 12);
                scop1.Text = litem.Text + "范围1: " + ds.Tables[0].Rows[j - 1]["LowValue"].ToString();
                //scop1.Font = new Font("宋体", 12);

                scop1.Name = litem.Text;
                locY1 -= 20;
                locX1 += 0;
                scop1.Location = new Point(locX1, locY1);
                this.Controls.Add(scop1);

                for (int i = 1; i <= int.Parse(ds.Tables[0].Rows[j - 1]["testvalueqty"].ToString()); i++)
                {
                    Label l = new Label();
                    t = new TextBox();
                    locX1 += 0;
                    counter01 += 1;
                    locY1 += l.Height;
                    l.Name = "lbl" + counter01.ToString();
                    l.Text = "测试值" + counter01.ToString();
                    l.Location = new Point(locX1, locY1);
                    this.panel1.Controls.Add(l);
                    locY1 += t.Height + 3;
                    t.Name = litem.Text + ":" + i.ToString();
                    t.Text = "";
                    t.Location = new Point(locX1, locY1);
                    this.panel1.Controls.Add(t);
                }
                counter01 = 0;
                locY1 = 0;
                locX1 += 160;
            }

        }
    }
}