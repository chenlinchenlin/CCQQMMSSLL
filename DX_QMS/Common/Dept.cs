using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace DX_QMS.Common
{
    class Dept
    {
        public static string SelectDeptNameById(string deptId)
        {
            SqlParameter[] deptPara = new SqlParameter[1];
            deptPara[0] = new SqlParameter("@deptid", deptId);

            return DbAccess.ExecuteScalar(CommandType.StoredProcedure, "dept_selectNameByID", deptPara).ToString();
        }

        //选择生产部门
        public static DataSet SelectProductDept()
        {
            return DbAccess.DataAdapterByCmd(CommandType.Text, "select deptid,deptname from dept where state='Y'", null);
        }

        //select all record
        public static DataSet SelectAll(string groupid, string deptid)
        {
            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@groupid", groupid);
            para[1] = new SqlParameter("@deptid", deptid);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "dept_SelectAll", para);
        }

        //add
        public static int AddRecordByKey(string deptId, string deptname, string state)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@deptid", deptId);
            para[1] = new SqlParameter("@deptname", deptname);
            para[2] = new SqlParameter("@state", state);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "dept_AddByKey", para);
        }

        //update
        public static int UpdateRecordByKey(string deptId, string newDeptName, string state)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@deptid", deptId);
            para[1] = new SqlParameter("@newdeptname", newDeptName);
            para[2] = new SqlParameter("@state", state);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "dept_UpdateByKey", para);
        }

        //delete
        public static int DeleteRecordByKey(string deptId)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@deptid", deptId);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "dept_DeleteByKey", para);
        }

        //select info by condition
        public static DataSet SelectInfoByCondition(string deptid, string deptname, string state)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@deptid", deptid);
            para[1] = new SqlParameter("@deptname", deptname);
            para[2] = new SqlParameter("@state", state);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Dept_SelectInfoByCondition", para);
        }
        public static int bluebooth(string sn)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@sn", sn);
            para[1] = new SqlParameter("@bluetooth_addr", SqlDbType.NVarChar, 50);
            para[1].Direction = ParameterDirection.Output;
            para[2] = new SqlParameter("@msg", SqlDbType.Int);
            para[2].Direction = ParameterDirection.Output;
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "BLUETOOTH_ASSIGN", para);
        }
    }
}
