using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace DX_QMS.Common
{
    class Users
    {
        public static DataSet SelectInfoById(string UserId)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@userid", UserId);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Users_SelectInfoByUserId", para);
        }

        //add
        public static int AddRecordByKey(string userId, string userName, string password, string groupId, string deptId, string tel, string deptid2, string groupBSid)
        {
            SqlParameter[] para = new SqlParameter[8];
            para[0] = new SqlParameter("@userid", userId);
            para[1] = new SqlParameter("@username", userName);
            para[2] = new SqlParameter("@password", password);
            para[3] = new SqlParameter("@groupid", groupId);
            para[4] = new SqlParameter("@deptid", deptId);
            para[5] = new SqlParameter("@tel", tel);
            para[6] = new SqlParameter("@deptid2", deptid2);
            para[7] = new SqlParameter("@groupBSid", groupBSid);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Users_AddRecordByKey", para);
        }

        //update
        public static int UpdateByKey(string userId, string userName, string groupId, string deptId, string tel, string deptid2, string groupBSid)
        {
            SqlParameter[] para = new SqlParameter[7];
            para[0] = new SqlParameter("@userid", userId);
            para[1] = new SqlParameter("@username", userName);
            para[2] = new SqlParameter("@groupid", groupId);
            para[3] = new SqlParameter("deptid", deptId);
            para[4] = new SqlParameter("@tel", tel);
            para[5] = new SqlParameter("@deptid2", deptid2);
            para[6] = new SqlParameter("@groupBSid", groupBSid);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Users_UpdateByKey", para);
        }

        //delete
        public static int DeleteByKey(string userId)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@userid", userId);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Users_DeleteByKey", para);
        }

        ///////////////////////////////////////////////////////
        ////     以下是自定义的方法
        //////////////////////////////////////////////////////

        //select userInfo by DeptId
        public static DataSet SelectUsersByGroupId(string groupId, string deptid)
        {
            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@groupid", groupId);
            para[1] = new SqlParameter("@deptid", deptid);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Users_SelectUsersByGroupId", para);
        }
        //update Password
        public static int UpdatePasswordByUserId(string userId, string oldPwd, string newPwd)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@userid", userId);
            para[1] = new SqlParameter("@oldpassword", oldPwd);
            para[2] = new SqlParameter("@newpassword", newPwd);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Users_UpdatePwdByUserId", para);
        }









        //clearPwd
        public static int ClearPassword(string userId, string password)
        {
            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@userid", userId);
            para[1] = new SqlParameter("@password", password);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Users_ClearPwd", para);
        }

        public static DataSet User_login(string userId, string password)
        {
            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@userid", userId);
            para[1] = new SqlParameter("@password", password);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Users_LoginIn", para);
        }

        public static DataSet QMS_User_login(string userId, string password)
        {
            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@userid", userId);
            para[1] = new SqlParameter("@password", password);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "QMS_Users_LoginIn", para);
        }


        //select info by condition
        public static DataSet SelectByConditon(string userid, string username, string deptid)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@userid", userid);
            para[1] = new SqlParameter("@username", username);
            para[2] = new SqlParameter("@deptid", deptid);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Users_SelectByCondition", para);
        }
        public static int AddTTSSet(string opertype, string ttype, string tdate, string time1, string time2, string tcontent, string maleorfemale, string forqty, string email, string userid)
        {
            SqlParameter[] para = new SqlParameter[10];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@ttype", ttype);
            para[2] = new SqlParameter("@tdate", tdate);
            para[3] = new SqlParameter("@time1", time1);
            para[4] = new SqlParameter("@time2", time2);
            para[5] = new SqlParameter("@tcontent", tcontent);
            para[6] = new SqlParameter("@maleorfemale", maleorfemale);
            para[7] = new SqlParameter("@forqty", forqty);
            para[8] = new SqlParameter("@email", email);
            para[9] = new SqlParameter("@userid", userid);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "SMT_TTSSetAdd", para);
        }
    }
}
