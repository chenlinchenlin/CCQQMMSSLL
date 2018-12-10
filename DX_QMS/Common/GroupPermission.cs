using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace DX_QMS.Common
{
    class GroupPermission
    {

        public static string  conversion(string permi)
        {
            string permission = "False";
            if (permi == null)
            {
                return permission;
            }
            else
            {
                return permi;
            }
        }
        public static bool AddNewRecordByGroupID(string groupID, string IsWeb)
        {
            bool temp = false;
            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@groupID", groupID);
            para[1] = new SqlParameter("@IsWeb", IsWeb);

            if (DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "GroupPermission_Add", para) > 0)
                temp = true;
            return temp;
        }

        //delete record by groupID
        public static bool DeleteRecordByGroupID(string groupID)
        {
            bool temp = false;
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@groupID", groupID);

            if (DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "GroupPermission_DeleteRecordByGroupID", para) > 0)
                temp = true;
            return temp;
        }

        //select module by grouID
        public static DataSet SelectModuleByGroupID(string groupID)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@groupID", groupID);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "GroupPermission_SelectModuleByGroupID", para);
        }

        //Select MainPageRule by groupID
        public static DataSet SelectRuleForMainForm(string groupID)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("groupID", groupID);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "GroupPermission_SelectRuleForMainForm", para);
        }

        //Select Rule by groupID & FormID
        public static Dictionary<string, bool> SelectRulesForForm(string moduleID, string groupID)
        {
            Dictionary<string, bool> dictionary = new Dictionary<string, bool>();

            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@moduleID", moduleID);
            para[1] = new SqlParameter("@groupID", groupID);

            DataTable dt = DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "GroupPermission_SelectRulesForForm", para).Tables[0];
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                    dictionary.Add(dt.Columns[i].ColumnName, (bool)dt.Rows[0][i]);
            }
            return dictionary;
        }

        public static Dictionary<string, bool> QMS_SelectRulesForForm(string post, string menuName)
        {
            Dictionary<string, bool> dictionary = new Dictionary<string, bool>();

            SqlParameter[] para = new SqlParameter[2];
            para[0] = new SqlParameter("@post", post);
            para[1] = new SqlParameter("@menuName", menuName);

            DataTable dt = DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "QMS_SelectRulesForForm", para).Tables[0];
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                    dictionary.Add(dt.Columns[i].ColumnName, (bool)dt.Rows[0][i]);
            }
            else
            {
                dictionary.Add("hasQuery", false);
                dictionary.Add("hasInsert", false);
                dictionary.Add("hasUpdate", false);
                dictionary.Add("hasDelete", false);
                dictionary.Add("hasPrint", false);
                dictionary.Add("hasAudit", false);    // hasQuery,hasInsert,hasUpdate,hasDelete,hasPrint,hasAudit
            }
            return dictionary;
        }





    }
}
