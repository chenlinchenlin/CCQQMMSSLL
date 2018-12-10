using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Collections;

namespace DX_QMS.Common
{
    class Module
    {
        public static DataSet SelectInfoByKey(string mID)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@mID", mID);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Module_SelectInfoByKey", para);
        }

        //select all information
        public static DataSet SelectAll()
        {
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Module_SelectAll", null);
        }

        //select all moduleID
        public static DataSet SelectAllModuleID(string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Module_SelectAllModuleID", para);
        }

        //add
        public static int AddRecordByKey(string mID, string mName, string mWindow, int mNo, string mParent, string mRule, string IsOper, string mDepts, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[9];
            para[0] = new SqlParameter("@mID", mID);
            para[1] = new SqlParameter("@mName", mName);
            para[2] = new SqlParameter("@mWindow", mWindow);
            para[3] = new SqlParameter("@mNo", mNo);
            para[4] = new SqlParameter("@mParent", mParent);
            para[5] = new SqlParameter("@mRule", mRule);
            para[6] = new SqlParameter("@IsOper", IsOper);
            para[7] = new SqlParameter("@mDepts", mDepts);
            para[8] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Module_AddRecordByKey", para);
        }

        //update
        public static int UpdateRecordByKey(string mID, string mName, string mWindow, int mNo, string mParent, string mRule, string IsOper, string mDepts, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[9];
            para[0] = new SqlParameter("@mID", mID);
            para[1] = new SqlParameter("@mName", mName);
            para[2] = new SqlParameter("@mWindow", mWindow);
            para[3] = new SqlParameter("@mNo", mNo);
            para[4] = new SqlParameter("@mParent", mParent);
            para[5] = new SqlParameter("@mRule", mRule);
            para[6] = new SqlParameter("@IsOper", IsOper);
            para[7] = new SqlParameter("@mDepts", mDepts);
            para[8] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Module_UpdateRecordByKey", para);
        }

        //delete
        public static int DeleteRecordByKey(string mID)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@mID", mID);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Module_DeleteRecordByKey", para);
        }

        //select moduleID by deptID
        public static ArrayList SelectMidByDept(string deptid, string IsWeb)
        {
            ArrayList tempList = new ArrayList();
            DataSet ds = SelectAllModuleID(IsWeb);
            if (ds.Tables[0].Rows.Count < 1)
                return null;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                if (ds.Tables[0].Rows[i][1].ToString().Contains(deptid))
                    tempList.Add(ds.Tables[0].Rows[i][0].ToString());
            }
            return tempList;
        }

        //select cmenu module 
        public static DataSet SelectMenuModule(string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@IsWeb", IsWeb);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Module_SelectMenuModule", para);
        }

        //select module by module left join on groupPemmision
        public static DataSet SelectModulesByPermission(string groupID, string deptid, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@groupID", groupID);
            para[1] = new SqlParameter("@deptid", deptid);
            para[2] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Module_SelectModulesByPermission", para);
        }

        //select info by condition
        public static DataSet SelectInfoByCondition(string mID, string mName, string IsWeb)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@mID", mID);
            para[1] = new SqlParameter("@mName", mName);
            para[2] = new SqlParameter("@IsWeb", IsWeb);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Module_SelectInfoByCondition", para);
        }
    }
}
