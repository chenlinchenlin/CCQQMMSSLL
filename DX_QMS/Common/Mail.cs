using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace DX_QMS.Common
{
    class Mail
    {
        public static string MailOperate(string action, string dept, string username, string userid, string type, string loginuser, string id, string mailType)
        {
            SqlParameter[] para = new SqlParameter[9];
            para[0] = new SqlParameter("@operType", action);
            para[1] = new SqlParameter("@deptid", dept);
            para[2] = new SqlParameter("@username", username);
            para[3] = new SqlParameter("@usermail", userid);
            para[4] = new SqlParameter("@sendtype", type);
            para[5] = new SqlParameter("@loginuser", loginuser);
            para[6] = new SqlParameter("@id", id);
            para[7] = new SqlParameter("@mailType", mailType);
            para[8] = new SqlParameter("@resultMsg", SqlDbType.VarChar, 200);
            para[8].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "SYS_ProdRepair_Mail", para);
            return para[8].Value.ToString();
        }


        public static DataSet MailOperate(string action, string dept, string username, string mailType)
        {
            SqlParameter[] para = new SqlParameter[9];
            para[0] = new SqlParameter("@operType", action);
            para[1] = new SqlParameter("@deptid", dept);
            para[2] = new SqlParameter("@username", username);
            para[3] = new SqlParameter("@usermail", "");
            para[4] = new SqlParameter("@sendtype", "");
            para[5] = new SqlParameter("@loginuser", "");
            para[6] = new SqlParameter("@id", "");
            para[7] = new SqlParameter("@mailType", mailType);
            para[8] = new SqlParameter("@resultMsg", SqlDbType.VarChar, 200);
            para[8].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "SYS_ProdRepair_Mail", para);
        }
    }
}
