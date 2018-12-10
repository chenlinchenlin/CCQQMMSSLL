using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class Login :DevExpress.XtraEditors.XtraForm
    {
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hwdn, int wMsg, int mParam, int lParam);
        public static string manager;
        public static string post;
        public static string groupId;
        public static string deptId;
        private bool pbIsVerify = false;
        public string pbVerifyUser = "";
        public static string userId;
        public static string username;
        private bool isStop = false;
        private byte[] bSoundStream;
        public Login()
        {
            InitializeComponent();
           // DevExpress.XtraEditors.XtraForm form = new DevExpress.XtraEditors.XtraForm();
            //this.Opacity = 0.8;
            //this.BackColor = Color.White;
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            this.BackColor = Color.White;
            this.Opacity = 0.9;
            //this.BackColor = Color.Transparent;
            //this.BackColor = Color.FromArgb(180, 40, 60, 82);
            // this.TransparencyKey = Color.White;// this.BackColor;// Color.White;
            //System.IO.MemoryStream ms = new System.IO.MemoryStream(barcode.Properties.Resources.DiamondBlue);
            //ms.Flush();
            //// this.skinEngine1.SkinStream = ms;
            //ms.Dispose();
        }

        public Login(string sLockUser)
        {
            // InitializeComponent();
            // isStop = true;
            // this.bSoundStream = new byte[barcode.Properties.Resources.SIREN3.Length];
            // barcode.Properties.Resources.SIREN3.Read(this.bSoundStream, 0, bSoundStream.Length);
            // System.IO.MemoryStream ms = new System.IO.MemoryStream(barcode.Properties.Resources.DiamondBlue);
            // ms.Flush();
            ////  this.skinEngine1.SkinStream = ms;
            // ms.Dispose();
            // this.pbIsVerify = true;
            // this.pbVerifyUser = sLockUser;
            // this.btnreset.Focus();
            // this.btnreset.Select();
        }

        public void FrmClickMeans(Form Frm_Tem, int n)
        {
            switch (n) 
            {
                case 0: 
                    Frm_Tem.WindowState = FormWindowState.Minimized;
                    break;
                case 1:	 
                    this.Close();
                    Application.Exit();
                    break;
            }
        }


        public static PictureBox Tem_PictB = new PictureBox();  
        public void ImageSwitch(object sender, int n, int ns)
        {
            Tem_PictB = (PictureBox)sender;
            switch (n)
            {
                case 0:
                    {
                        Tem_PictB.Image = null;
                        if (ns == 0)
                            Tem_PictB.Image = Properties.Resources.zuixiao1;
                            Tem_PictB.SizeMode = PictureBoxSizeMode.StretchImage;
                        if (ns == 1)
                            Tem_PictB.Image = Properties.Resources.zuixiao;
                        break;
                    }
                case 1:
                    {
                        Tem_PictB.Image = null;
                        if (ns == 0)
                            Tem_PictB.Image = Properties.Resources.guanbi1;
                            Tem_PictB.SizeMode = PictureBoxSizeMode.StretchImage;
                        if (ns == 1)
                            Tem_PictB.Image = Properties.Resources.guanbi;
                        break;
                    }
            }
        }

        private void pictureBox_min_Click(object sender, EventArgs e)
        {
            FrmClickMeans(this, Convert.ToInt16(((PictureBox)sender).Tag.ToString()));
        }

        private void pictureBox_min_MouseEnter(object sender, EventArgs e)
        {
            ImageSwitch(sender, Convert.ToInt16(((PictureBox)sender).Tag.ToString()), 0);
        }

        private void pictureBox_min_MouseLeave(object sender, EventArgs e)
        {
            ImageSwitch(sender, Convert.ToInt16(((PictureBox)sender).Tag.ToString()), 1);
        }

        public void CreateNode(XmlDocument xmlDoc, XmlNode parentNode, string name, string value)
        {
            XmlNode node = xmlDoc.CreateNode(XmlNodeType.Element, name, null);
            node.InnerText = value;
            parentNode.AppendChild(node);
        }
        public void CreateXmlFile()
        {
            XmlDocument xmlDoc = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            //创建根节点
            XmlNode root = xmlDoc.CreateElement("root");
            xmlDoc.AppendChild(root);
            XmlNode user = xmlDoc.CreateElement("user");
            root.AppendChild(user);
            CreateNode(xmlDoc, user, "id", txtuser.Text);
            CreateNode(xmlDoc, user, "pwd", txtpassword.Text);
            try
            {
                xmlDoc.Save(Application.StartupPath + "\\.xml");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void Login_Form_Load(object sender, EventArgs e)
        {
            pictureBox_min.Image = null;//清空PictuteBox控件
            pictureBox_min.Image = Properties.Resources.zuixiao;//显示最小化按钮的图片
            pictureBox_min.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox_max.Image = null;//清空PictuteBox控件
            pictureBox_max.Image = Properties.Resources.guanbi;//显示关闭按钮的图片
            pictureBox_max.SizeMode = PictureBoxSizeMode.StretchImage;

            if (!File.Exists(Application.StartupPath + "\\.xml"))
            {
                CreateXmlFile();
                return;
            }

            else
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Application.StartupPath + "\\.xml");
                XmlNodeList userNodes = xmlDoc.SelectSingleNode("root").ChildNodes;
                foreach (XmlNode userNode in userNodes)
                {
                    txtuser.Properties.Items.Add(userNode.FirstChild.InnerText);                   
                }

            }
        }

        private void Login_Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!this.pbIsVerify)
            {
                Application.Exit();
            }
        }

        private void Login_Form_Shown(object sender, EventArgs e)
        {
            while (isStop)
            {
                System.Windows.Forms.Application.DoEvents();
                //  sound.PlayByStream(this.bSoundStream);
            }
        }

        public string GetPwd(string userId)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(Application.StartupPath + "\\.xml");
            XmlNodeList userNodes = xmlDoc.SelectSingleNode("root").ChildNodes;
            foreach (XmlNode user in userNodes)
            {
                XmlNodeList idandpwd = user.ChildNodes;
                foreach (XmlNode idnode in idandpwd)
                {
                    if (idnode.InnerText == txtuser.Text)
                    {
                        string pwd = idnode.NextSibling.InnerText;
                        if (pwd == "")
                        {
                            return "";
                        }
                        return pwd;
                    }
                }
            }
            return "";
        }
        private void txtuser_SelectedValueChanged(object sender, EventArgs e)  
        {
            txtpassword.Text = GetPwd(txtuser.Text);   
            if (txtpassword.Text != "")
            {
                cboxRemember.Checked = true;
            }
            else
            {
                cboxRemember.Checked = false;
            }
        }

        private void btnlogin_Click(object sender, EventArgs e)
        {
            if (txtuser.Text.Trim() == "" || txtpassword.Text.Trim() == "") return;
            if (this.pbIsVerify && txtuser.Text.Trim().Equals(this.pbVerifyUser))
            {
                txtuser.Focus();
                txtuser.SelectAll();
                MessageBox.Show("不能用锁住的用户解锁,请重新输入!", "提示!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataSet ds = Users.QMS_User_login(txtuser.Text.Trim(), DbAccess.Encrypt(txtpassword.Text.Trim()));

            if (ds.Tables.Count > 0)  
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (this.pbIsVerify)
                    {
                        pbVerifyUser = txtuser.Text.Trim();
                        this.DialogResult = DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        userId = txtuser.Text.Trim();
                        manager = ds.Tables[0].Rows[0]["manager"].ToString().Trim();
                        post = ds.Tables[0].Rows[0]["post"].ToString().Trim();
                        username = ds.Tables[0].Rows[0]["userName"].ToString().Trim();
                       ////// deptId2 = ds.Tables[0].Rows[0]["deptid2"].ToString().Trim();
                        SetUser();
                        Form_QMS MF = new Form_QMS();
                        this.Hide();
                        login_loading.ShowWaitForm();
                        MF.WindowState = FormWindowState.Maximized;
                        MF.Show();
                        login_loading.CloseWaitForm();
                    }
                }
                else
                {
                    txtpassword.Focus();
                    txtpassword.SelectAll();
                    MessageBox.Show("密码错误,请重新输入!", "提示!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                txtuser.Focus();
                txtuser.SelectAll();
                MessageBox.Show("用户名不存在,请重新输入!", "提示!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SetUser()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(Application.StartupPath + "\\.xml");
            XmlNodeList userNodes = xmlDoc.SelectSingleNode("root").ChildNodes;
            foreach (XmlNode userNode in userNodes)
            {
                XmlNodeList idandpwd = userNode.ChildNodes;
                foreach (XmlNode idnode in idandpwd)
                {
                    if (idnode.InnerText == txtuser.Text)
                    {
                        if (cboxRemember.Checked)
                        {
                            idnode.NextSibling.InnerText = txtpassword.Text;
                            xmlDoc.Save(Application.StartupPath + "\\.xml");
                            return;
                        }
                        else
                        {
                            idnode.NextSibling.InnerText = "";
                            xmlDoc.Save(Application.StartupPath + "\\.xml");
                            return;
                        }
                    }
                }
            }
            XmlNode root = xmlDoc.SelectSingleNode("root");
            XmlElement user = xmlDoc.CreateElement("user");
            XmlElement id = xmlDoc.CreateElement("id");
            XmlElement pwd = xmlDoc.CreateElement("pwd");
            id.InnerText = txtuser.Text;
            user.AppendChild(id);
            if (cboxRemember.Checked)
            {
                pwd.InnerText = txtpassword.Text;
            }
            else
            {
                pwd.InnerText = "";
            }
            user.AppendChild(pwd);
            root.AppendChild(user);
            xmlDoc.Save(Application.StartupPath + "\\.xml");

        }  
        private void btnreset_Click(object sender, EventArgs e)
        {
            txtuser.Text = "";
            txtpassword.Text = "";
            if (!this.pbIsVerify)
                txtuser.Focus();
        }
        private void Login_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();//用来释放被当前线程中某个窗口捕获的光标
            SendMessage(this.Handle, 0x0112, 0xF010 + 0x0002, 0);//向Windows发送拖动窗体的消息
        }

        private void txtpassword_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtpassword.Text != "")
            {
                btnlogin_Click(sender, e);
            }
        }
    }
}
