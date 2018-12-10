using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace DX_QMS.Common
{
    class Groups
    {
        public static DataSet SelectAll(string groupid, string deptid, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@groupid", groupid);
            para[1] = new SqlParameter("@deptid", deptid);
            para[2] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Groups_SelectAllRecord", para);
        }

        //select GroupName by GroupId
        public static string SelectNameById(string groupId)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@groupid", groupId);

            return DbAccess.ExecuteScalar(CommandType.StoredProcedure, "Groups_SelectNameById", para).ToString();
        }

        //select dept by Groupid
        public static string SelectDeptByGroupid(string groupId)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@groupid", groupId);

            return DbAccess.ExecuteScalar(CommandType.StoredProcedure, "Groups_SelectDeptByGroupid", para).ToString();
        }

        //insert into one record
        public static int InsertOneRecord(string groupid, string groupname, string remark, string deptid, string creatGroup, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[6];
            para[0] = new SqlParameter("@groupid", groupid);
            para[1] = new SqlParameter("@groupname", groupname);
            para[2] = new SqlParameter("@remark", remark);
            para[3] = new SqlParameter("@deptid", deptid);
            para[4] = new SqlParameter("@creatGroup", creatGroup);
            para[5] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Groups_InsertRecordById", para);
        }

        //update record by key
        public static int UpdateRecordByKey(string groupid, string groupname, string remark, string deptid, string creatGroup, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[6];
            para[0] = new SqlParameter("@groupid", groupid);
            para[1] = new SqlParameter("@groupname", groupname);
            para[2] = new SqlParameter("@remark", remark);
            para[3] = new SqlParameter("@deptid", deptid);
            para[4] = new SqlParameter("@creatGroup", creatGroup);
            para[5] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Groups_UpdateRecordByKey", para);
        }

        //delete record by key
        public static int DeleteRecordByKey(string groupid)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@groupid", groupid);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Groups_DeleteRecordByKey", para);
        }

        //Select info by condition
        public static DataSet SelectByCondition(string groupid, string groupname, string deptid, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[4];
            para[0] = new SqlParameter("@groupid", groupid);
            para[1] = new SqlParameter("@groupname", groupname);
            para[2] = new SqlParameter("@deptid", deptid);
            para[3] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Groups_SelectByCondition", para);
        }
    }
}
