using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Net.Mail;

namespace DX_QMS.Common
{
    class ProdTest
    {
        SqlTransaction tran = null;

        public bool DeleteSAP_serialcode(string dept, string workno, string lotno, int qty, string countrytype, string userid)
        {
            SqlParameter[] para = new SqlParameter[6];
            para[0] = new SqlParameter("@dept", dept);
            para[1] = new SqlParameter("@workno", workno);
            para[2] = new SqlParameter("@lotno", lotno);
            para[3] = new SqlParameter("@qty", qty);
            para[4] = new SqlParameter("@countrytype", countrytype);
            para[5] = new SqlParameter("@userid", userid);
            return DbAccess.ExecuteNonQueryByTranNew(CommandType.StoredProcedure, "SAP_serialcodedel", para, tran);
        }

        public static string InsertProdTest(string sn, string wo, string mo, string model, string shipaddress, string checkuser, string checkdate, string checkresult, string remark, string loginuser)
        {
            SqlParameter[] para = new SqlParameter[11];
            para[0] = new SqlParameter("@sn", sn);
            para[1] = new SqlParameter("@wo", wo);
            para[2] = new SqlParameter("@mo", mo);
            para[3] = new SqlParameter("@model", model);
            para[4] = new SqlParameter("@shipaddress", shipaddress);
            para[5] = new SqlParameter("@checkuser", checkuser);
            para[6] = new SqlParameter("@checkdate", checkdate);
            para[7] = new SqlParameter("@checkresult", checkresult);
            para[8] = new SqlParameter("@remarks", remark);
            para[9] = new SqlParameter("@operUser", loginuser);
            para[10] = new SqlParameter("@errormsg", SqlDbType.VarChar, 200);
            para[10].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "SYS_ProdTest_Add", para);
            return para[10].Value.ToString();

        }

        public static string UpdateProdTest(string sn, string wo, string mo, string model, string shipaddress, string checkuser, string checkdate, string checkresult, string remark, string loginuser)
        {
            SqlParameter[] para = new SqlParameter[11];
            para[0] = new SqlParameter("@sn", sn);
            para[1] = new SqlParameter("@wo", wo);
            para[2] = new SqlParameter("@mo", mo);
            para[3] = new SqlParameter("@model", model);
            para[4] = new SqlParameter("@shipaddress", shipaddress);
            para[5] = new SqlParameter("@checkuser", checkuser);
            para[6] = new SqlParameter("@checkdate", checkdate);
            para[7] = new SqlParameter("@checkresult", checkresult);
            para[8] = new SqlParameter("@remarks", remark);
            para[9] = new SqlParameter("@operUser", loginuser);
            para[10] = new SqlParameter("@errormsg", SqlDbType.VarChar, 200);
            para[10].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "SYS_ProdTest_Update", para);
            return para[10].Value.ToString();

        }

        public static int DeleteProdTest(string sn)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@sn", sn);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "SYS_ProdTest_Delete", para);
        }


        public static DataSet GetProdTestInfo(string sn, string wo, string mo, string model, string shipaddress, string checkuser, string startcheckdate, string checkresult, string endcheckdate)
        {
            SqlParameter[] para = new SqlParameter[9];
            para[0] = new SqlParameter("@sn", sn);
            para[1] = new SqlParameter("@wo", wo);
            para[2] = new SqlParameter("@mo", mo);
            para[3] = new SqlParameter("@model", model);
            para[4] = new SqlParameter("@shipaddress", shipaddress);
            para[5] = new SqlParameter("@checkuser", checkuser);
            para[6] = new SqlParameter("@startcheckdate", startcheckdate);
            para[7] = new SqlParameter("@checkresult", checkresult);
            para[8] = new SqlParameter("@endcheckdate", endcheckdate);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "SYS_ProdTest_Select", para);
        }

        public static void SendMail(string dept, string subject, string body, string mailGroup)
        {
            try
            {
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Host = "smtp.hytera.com";
                SmtpClient.Port = 25;
                SmtpClient.Timeout = 60;
                SmtpClient.Credentials = new System.Net.NetworkCredential("publicpostman@hytera.com", "publicpostman");

                MailMessage MailMessage_Mai = new MailMessage();
                MailMessage_Mai.To.Clear();

                DataTable temp = Mail.MailOperate("query", dept, "", mailGroup).Tables[0];
                if (temp.Rows.Count == 0)
                    return;
                for (int i = 0; i < temp.Rows.Count; i++)
                {
                    if ("to".Equals(temp.Rows[i]["sendType"].ToString()))
                        MailMessage_Mai.To.Add(new MailAddress(temp.Rows[i]["usermail"].ToString() + "@hytera.com"));
                    else
                        MailMessage_Mai.CC.Add(new MailAddress(temp.Rows[i]["usermail"].ToString() + "@hytera.com"));
                }
                MailMessage_Mai.From = new MailAddress("publicpostman@hytera.com", "QMS系统");
                MailMessage_Mai.Subject = subject;
                MailMessage_Mai.SubjectEncoding = System.Text.Encoding.UTF8;
                //邮件正文                
                MailMessage_Mai.Body = body;
                MailMessage_Mai.BodyEncoding = System.Text.Encoding.UTF8;
                object e = new object();
                SmtpClient.SendAsync(MailMessage_Mai, e);
            }
            catch (Exception e)
            {
            }
        }

        public static void SendMailSysPlan(string dept, string subject, string body, string mailGroup)
        {
            try
            {
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Host = "smtp.hytera.com";
                SmtpClient.Port = 25;
                SmtpClient.Timeout = 60;
                SmtpClient.Credentials = new System.Net.NetworkCredential("publicpostman@hytera.com", "publicpostman");

                MailMessage MailMessage_Mai = new MailMessage();
                MailMessage_Mai.To.Clear();

                DataTable temp = Mail.MailOperate("query", dept, "", mailGroup).Tables[0];
                if (temp.Rows.Count == 0)
                    return;
                for (int i = 0; i < temp.Rows.Count; i++)
                {
                    if ("to".Equals(temp.Rows[i]["sendType"].ToString()))
                        MailMessage_Mai.To.Add(new MailAddress(temp.Rows[i]["usermail"].ToString()));
                    else
                        MailMessage_Mai.CC.Add(new MailAddress(temp.Rows[i]["usermail"].ToString()));
                }
                MailMessage_Mai.From = new MailAddress("publicpostman@hytera.com", "QMS系统");
                MailMessage_Mai.Subject = subject;
                MailMessage_Mai.SubjectEncoding = System.Text.Encoding.UTF8;

                //邮件正文                
                MailMessage_Mai.Body = body;
                MailMessage_Mai.BodyEncoding = System.Text.Encoding.UTF8;
                MailMessage_Mai.IsBodyHtml = true;//是否是HTML邮件 
                MailMessage_Mai.Priority = MailPriority.High;

                object e = new object();
                SmtpClient.SendAsync(MailMessage_Mai, e);
            }
            catch (Exception e)
            {
                //throw new Exception(e.Message);
            }
        }

        public static void SendMailToSQE(string dept, string subject, string body, string mailGroup, string username)
        {
            try
            {
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Host = "emaillg.hytera.com";   //  mail.hytera.com 
                SmtpClient.Port = 587;
                SmtpClient.Timeout = 60;
                SmtpClient.Credentials = new System.Net.NetworkCredential("x90171", "Password@1");

                MailMessage MailMessage_Mai = new MailMessage();
                MailMessage_Mai.To.Clear();

                foreach (string s in username.Split(','))
                {
                    DataTable temp = Mail.MailOperate("query", dept, s, mailGroup).Tables[0];
                    if (temp.Rows.Count == 0)
                        return;
                    for (int i = 0; i < temp.Rows.Count; i++)
                    {
                        if ("to".Equals(temp.Rows[i]["sendType"].ToString()))
                            MailMessage_Mai.To.Add(new MailAddress(temp.Rows[i]["usermail"].ToString()));
                        else
                            MailMessage_Mai.CC.Add(new MailAddress(temp.Rows[i]["usermail"].ToString()));
                    }
                }
                MailMessage_Mai.From = new MailAddress("public.postman@hytera.com", "QMS系统");
                MailMessage_Mai.Subject = subject;
                MailMessage_Mai.SubjectEncoding = System.Text.Encoding.UTF8;
                //邮件正文                
                MailMessage_Mai.Body = body;
                MailMessage_Mai.BodyEncoding = System.Text.Encoding.UTF8;
                object e = new object();
                SmtpClient.SendAsync(MailMessage_Mai, e);
                MessageBox.Show("邮件发送完成!");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }


        public static void SendMail(string to,string subject, string body)
        {
            try
            {
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Host = "emaillg.hytera.com";
                SmtpClient.Port = 587;
                SmtpClient.Timeout = 60;
                SmtpClient.Credentials = new System.Net.NetworkCredential("x90171", "Password@1");

                MailMessage MailMessage_Mai = new MailMessage();


                MailMessage_Mai.To.Clear();

                MailMessage_Mai.To.Add(new MailAddress(to));

                MailMessage_Mai.From = new MailAddress("public.postman@hytera.com", "QMS系统");
                MailMessage_Mai.Subject = subject;
                MailMessage_Mai.SubjectEncoding = System.Text.Encoding.UTF8;
                //邮件正文                
                MailMessage_Mai.Body = body;
                MailMessage_Mai.BodyEncoding = System.Text.Encoding.UTF8;
                object e = new object();
                SmtpClient.SendAsync(MailMessage_Mai, e);

                MessageBox.Show("发送成功","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);

            }
            catch (Exception e)
            {
            }
        }



        public static void SendHTMLboyMail(string to, string subject, string body)
        {
            try
            {
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Host = "emaillg.hytera.com";
                SmtpClient.Port = 587;
                SmtpClient.Timeout = 60;
                SmtpClient.Credentials = new System.Net.NetworkCredential("x90171", "Password@1");

                MailMessage MailMessage_Mai = new MailMessage();


                MailMessage_Mai.To.Clear();

                MailMessage_Mai.To.Add(new MailAddress(to));

                MailMessage_Mai.From = new MailAddress("public.postman@hytera.com", "QMS系统");
                MailMessage_Mai.Subject = subject;
                MailMessage_Mai.SubjectEncoding = System.Text.Encoding.UTF8;
                //邮件正文                
                MailMessage_Mai.Body = body;
                MailMessage_Mai.IsBodyHtml = true;
                MailMessage_Mai.BodyEncoding = System.Text.Encoding.UTF8;
                object e = new object();
                SmtpClient.SendAsync(MailMessage_Mai, e);

                MessageBox.Show("发送成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception e)
            {
            }
        }

        public static void SendMailComponent(string dept, string subject, string body, string mailGroup, string attachFile)
        {
            try
            {
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Host = "smtp.hytera.com";
                SmtpClient.Port = 25;
                SmtpClient.Timeout = 60;
                SmtpClient.Credentials = new System.Net.NetworkCredential("publicpostman@hytera.com", "publicpostman");

                MailMessage MailMessage_Mai = new MailMessage();
                MailMessage_Mai.To.Clear();

                DataTable temp = Mail.MailOperate("query", dept, "", mailGroup).Tables[0];
                if (temp.Rows.Count == 0)
                    return;
                for (int i = 0; i < temp.Rows.Count; i++)
                {
                    if ("to".Equals(temp.Rows[i]["sendType"].ToString()))
                        MailMessage_Mai.To.Add(new MailAddress(temp.Rows[i]["usermail"].ToString()));
                    else
                        MailMessage_Mai.CC.Add(new MailAddress(temp.Rows[i]["usermail"].ToString()));
                }
                MailMessage_Mai.From = new MailAddress("publicpostman@hytera.com", "QMS系统");
                MailMessage_Mai.Subject = subject;
                MailMessage_Mai.SubjectEncoding = System.Text.Encoding.UTF8;

                //邮件正文                
                MailMessage_Mai.Body = body;
                MailMessage_Mai.BodyEncoding = System.Text.Encoding.UTF8;
                MailMessage_Mai.IsBodyHtml = true;//是否是HTML邮件 
                MailMessage_Mai.Priority = MailPriority.High;

                System.Net.Mail.Attachment myAttachment = new Attachment(attachFile);
                MailMessage_Mai.Attachments.Add(myAttachment);

                object e = new object();
                SmtpClient.SendAsync(MailMessage_Mai, e);
            }
            catch (Exception e)
            {
                //throw new Exception(e.Message);
            }
        }
        public static string HtmlBoy(DataTable dt ,string title,string content, string tablecontent)
        {
            string Htmlstr = " <!DOCTYPE html>  <html lang=\"en\">  <head> <meta charset=\"UTF-8\">  	<meta name=\"viewport\" content=\"width = device-width, initial-scale = 1.0\">  <meta http-equiv=\"X-UA-Compatible\" content=\"ie=edge\"> ";
            Htmlstr += " <title>"+ title + "</title>  <style>   .mail-desc{  padding: 10px 0;  font-size:14.0pt; font-family:\"宋体\";  color:#1F4E79;  mso-style-textfill-fill-color:#1F4E79;  mso-style-textfill-fill-alpha:100.0%; }  .table-desc{  width:716.0pt;  border:solid windowtext 1.0pt;  padding:0cm 5.4pt 0cm 5.4pt; height:29.25pt; } ";
            Htmlstr += " .table-title{  width:89.5pt;  border:solid windowtext 1.0pt;  border-top:none;background:#00B050;padding:0cm 5.4pt 0cm 5.4pt;  height:28.5pt;} .no-border{border-top:none;border-left:none; border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;}";
            Htmlstr += "  .MsoNormalTable{ width: 716.0pt; margin-left: -.15pt;  border-collapse: collapse; } .MsoNormal{ text-align: center; }  .MsoNormal .txt{  font-size:12.0pt;  color:black; } .table-content{  width:89.5pt;  border:solid windowtext 1.0pt;  border-top:none; background:#FFFFFF; padding:0cm 5.4pt 0cm 5.4pt; height:16.5pt;  } ";
            Htmlstr += "  .table-content .MsoNormal .txt{  font-size:10.0pt; font-family: \"微软雅黑\", sans-serif;  }   .bg-red{ background: red; } .font-size-11{  font-size:11.0pt; } </style>  </head>  ";
            Htmlstr += " <body>  <p class=\"mail-desc\">"+content+ "<span lang=\"EN-US\"></span></p>  	<table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"955\">  <tbody>  <tr>  <td width=\"955\" colspan=\"8\" class=\"table-desc\">  <p class=\"MsoNormal\"><span class=\"txt font-size-11\">"+tablecontent+ "<span lang=\"EN-US\">  </span></span></p> 	  ";
            Htmlstr += "  </td>  </tr>  ";

            string Htmlcol = "  <tr> ";
            string Htmlcolname = " ";
            int colcount = dt.Columns.Count;
            for (int col = 0;col < colcount; col++)
            {                               
                Htmlcolname += "  <td class=\"table-title\">  <p class=\"MsoNormal\"><span class=\"txt\">";
                Htmlcolname += dt.Columns[col].ColumnName;
                Htmlcolname += "</span></p>  </td>   ";
            }
            Htmlcol += Htmlcolname;
            Htmlcol += "  </tr>  ";
            Htmlstr += Htmlcol;

            string Htmlrow = "  ";           
            for (int row = 0;row<dt.Rows.Count;row ++)
            {
                Htmlrow = "   <tr>  ";
                string Htmlrowboy = " ";
                for (int rowcol = 0; rowcol < colcount; rowcol++)
                {
                    Htmlrowboy += " <td class=\"table-content\">   <p class=\"MsoNormal\"><span class=\"txt\">";
                    Htmlrowboy += dt.Rows[row][rowcol].ToString();
                    Htmlrowboy += "</span></p>  </td> ";
                }
                Htmlrow += Htmlrowboy;
                Htmlrow += "  </tr>  ";
                Htmlstr += Htmlrow;
            }
            Htmlstr += "   </tbody>    </table>  </body>  </html>  ";
            return Htmlstr;
        }


        public static string FitHtmlBoy(DataTable dt,string title, string content, string tablecontent)
        {
            int tablecolcount = dt.Columns.Count;
            int length = tablecolcount * 120;
            string Htmlstr = " <!DOCTYPE html>  <html lang=\"en\">  <head> <meta charset=\"UTF-8\">  	<meta name=\"viewport\" content=\"width = device-width, initial-scale = 1.0\">  <meta http-equiv=\"X-UA-Compatible\" content=\"ie=edge\"> ";
            Htmlstr += " <title>" + title + "</title>  <style>   .mail-desc{  padding: 10px 0;  font-size:14.0pt; font-family:\"宋体\";  color:#1F4E79;  mso-style-textfill-fill-color:#1F4E79;  mso-style-textfill-fill-alpha:100.0%; }  .table-desc{  width:716.0pt;  border:solid windowtext 1.0pt;  padding:0cm 5.4pt 0cm 5.4pt; height:29.25pt; } ";
            Htmlstr += " .table-title{  width:89.5pt;  border:solid windowtext 1.0pt;  border-top:none;background:#00B050;padding:0cm 5.4pt 0cm 5.4pt;  height:28.5pt;} .no-border{border-top:none;border-left:none; border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;}";
            Htmlstr += "  .MsoNormalTable{ width: 716.0pt; margin-left: -.15pt;  border-collapse: collapse; } .MsoNormal{ text-align: center; }  .MsoNormal .txt{  font-size:12.0pt;  color:black; } .table-content{  width:89.5pt;  border:solid windowtext 1.0pt;  border-top:none; background:#FFFFFF; padding:0cm 5.4pt 0cm 5.4pt; height:16.5pt;  } ";
            Htmlstr += "  .table-content .MsoNormal .txt{  font-size:10.0pt; font-family: \"微软雅黑\", sans-serif;  }   .bg-red{ background: red; } .font-size-11{  font-size:11.0pt; } </style>  </head>  ";
            Htmlstr += " <body>  <p class=\"mail-desc\">" + content + "<span lang=\"EN-US\"></span></p>  	<table class=\"MsoNormalTable\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\""+length.ToString()+"\">  <tbody>  <tr>  <td width=\""+length.ToString()+"\" colspan=\""+tablecolcount.ToString()+"\" class=\"table-desc\">  <p class=\"MsoNormal\"><span class=\"txt font-size-11\">" + tablecontent + "<span lang=\"EN-US\">  </span></span></p> 	  ";
            Htmlstr += "  </td>  </tr>  ";

            string Htmlcol = "  <tr> ";
            string Htmlcolname = " ";
            int colcount = dt.Columns.Count;
            for (int col = 0; col < colcount; col++)
            {
                Htmlcolname += "  <td class=\"table-title\">  <p class=\"MsoNormal\"><span class=\"txt\">";
                Htmlcolname += dt.Columns[col].ColumnName;
                Htmlcolname += "</span></p>  </td>   ";
            }
            Htmlcol += Htmlcolname;
            Htmlcol += "  </tr>  ";
            Htmlstr += Htmlcol;

            string Htmlrow = "  ";
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                Htmlrow = "   <tr>  ";
                string Htmlrowboy = " ";
                for (int rowcol = 0; rowcol < colcount; rowcol++)
                {
                    Htmlrowboy += " <td class=\"table-content\">   <p class=\"MsoNormal\"><span class=\"txt\">";
                    Htmlrowboy += dt.Rows[row][rowcol].ToString();
                    Htmlrowboy += "</span></p>  </td> ";
                }
                Htmlrow += Htmlrowboy;
                Htmlrow += "  </tr>  ";
                Htmlstr += Htmlrow;
            }
            Htmlstr += "   </tbody>    </table>  </body>  </html>  ";
            return Htmlstr;
        }



    }
}
