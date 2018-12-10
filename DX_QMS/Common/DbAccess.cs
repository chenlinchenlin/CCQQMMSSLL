using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.Collections;
using System.Security.Cryptography;
using DevExpress.XtraEditors.Controls;
using System.Windows.Forms;
using DevExpress.XtraGrid;

namespace DX_QMS.Common
{
    class DbAccess
    {
       // private static string connSql = "10.102.0.12";
       // public const string connSql = "server=192.168.0.176;database=BarcodeNew;user id=mesproxy;password=I2007pledge07my30life2011and05honor01to2014the09Night01Watch;Pooling=false";
         public const string connSql = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
        public const string conndjiSql = "server=10.100.0.128;database=DJIInterface;user id=dji;Pooling=false;password=dji2018";
        // public const string connSql = "server=192.168.0.204;database=BarcodeNew;user id=sa;Pooling=false;password=The0more7people0you7love3the7weaker8you8are";
        //public const string connSql = "server=SVN_PDMR;database=THDZ;user id=sa;password=hytera.thdz;Pooling=false";
        //public const string connSql = "server=192.168.0.176;database=THDZ;user id=sa;password=123456;Pooling=false";
        //2014年09月11日修改,只用于IQC检验入库接口  
        //public const string connOral = "Data Source=  (DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.104)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = PROD)));User ID=BARCODE;Password=BARCODE;Unicode=True";
        //public const string connOral = "Data Source=  (DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.2.50)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = PROD)));User ID=BARCODE;Password=BARCODE;Unicode=True";
        //public const string connOral = "Data Source=  (DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.106)(PORT = 1523)))(CONNECT_DATA =(SERVICE_NAME = TEST2)));User ID=BARCODE;Password=BARCODE;Unicode=True";
        //public const string connOral = "Data Source=  (DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.2.50)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = PROD)));User ID=BARCODE;Password=BARCODE;Unicode=True";
        public const string connOral = "Data Source=  (DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.2.50)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = PROD)));User ID=BARCODE;Password=BARCODE;";

        //public static SqlConnection sqlconn = new SqlConnection(connSql);
        //public static OracleConnection oraconn = new OracleConnection(connOral);

        //static DbAccess()
        //{
        //    try
        //    {
        //        ConfigurationManager.RefreshSection("connectionString");
        //        string[] arrEnCryto = ConfigurationManager.AppSettings.GetValues("connectionString");
        //        if (arrEnCryto != null && arrEnCryto.Length > 0)
        //            connSql = barclass.utility.Decrypt(arrEnCryto[0]);
        //    }
        //    catch { }
        //}

        public DbAccess()
        {
            SqlConnection.ClearAllPools();
            //OracleConnection.ClearAllPools();
        }
        /// <summary>
        /// 判断sql记录是否有存在,查询结果是否有记录
        /// </summary>
        /// <param name="strsql">SQL语句</param>
        /// <returns></returns>
        public static bool Exists(string StrSql)
        {
            object obj = SelectOneRecord(StrSql);
            int cmdresult;
            if ((object.Equals(obj, null)) || (object.Equals(obj, System.DBNull.Value)))
            {
                cmdresult = 0;
            }
            else
            {
                cmdresult = int.Parse(obj.ToString());
            }
            if (cmdresult == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        /// <summary>
        /// 从Oracle数据库查询数据，返回类型为DateSet
        /// </summary>
        /// <param name="oraStr">Oracle查询字符串</param>
        /// <returns></returns>
        /*public static DataSet SelectByOracle(string oraStr)
        {
            //select info by condition
            SqlParameter[] para = new SqlParameter[1];
            oraStr = oraStr.Replace("'", "''");
            para[0] = new SqlParameter("@oraStr", oraStr);

            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "SelectByOracle", para);

            //2010年11月9日修改
            //OracleConnection oraconn = new OracleConnection(connOral);
            //try
            //{
            //    if (oraconn.State != ConnectionState.Open)
            //        oraconn.Open();
            //    OracleDataAdapter da = new OracleDataAdapter(oraStr, oraconn);
            //    DataSet ds = new DataSet();
            //    da.Fill(ds);
            //    return ds;
            //}
            //catch(Exception ex)
            //{
            //    throw new Exception(ex.Message);
            //}
            //finally
            //{
            //    oraconn.Close();
            //}
        }
        */
        public static DataSet SelectByOracle(string oraStr)
        {
            //2010年11月9日修改

            Oracle.ManagedDataAccess.Client.OracleConnection oraconn = new Oracle.ManagedDataAccess.Client.OracleConnection(connOral);
            try
            {
                if (oraconn.State != ConnectionState.Open)
                    oraconn.Open();
                Oracle.ManagedDataAccess.Client.OracleDataAdapter da = new Oracle.ManagedDataAccess.Client.OracleDataAdapter(oraStr, oraconn);
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oraconn.Close();
            }
        }
        /// <summary>
        /// 从Sql数据库查询数据，返回类型为DateSet
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static DataSet SelectBySql(string str)
        {
            SqlConnection sqlconn = new SqlConnection(connSql);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();
                SqlDataAdapter da = new SqlDataAdapter(str, sqlconn);
                da.SelectCommand.CommandTimeout = 0;
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                sqlconn.Close();
            }
        }


        public static DataSet SelectBydjiSql(string str)
        {
            SqlConnection sqlconn = new SqlConnection(conndjiSql);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();
                SqlDataAdapter da = new SqlDataAdapter(str, sqlconn);
                da.SelectCommand.CommandTimeout = 0;
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                sqlconn.Close();
            }
        }


        /// <summary>
        /// 对Sql执行insert,update,delete操作，返回类型为bool
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool ExecuteSql(string str)
        {
            bool va = false;
            SqlConnection sqlconn = new SqlConnection(connSql);
            SqlCommand cmd = new SqlCommand(str, sqlconn);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();
                int temp = cmd.ExecuteNonQuery();
                if (temp > 0)
                    va = true;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                sqlconn.Close();
            }
            return va;
        }

        public static bool ExecutedjiSql(string str)
        {
            bool va = false;
            SqlConnection sqlconn = new SqlConnection(conndjiSql);
            SqlCommand cmd = new SqlCommand(str, sqlconn);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();
                int temp = cmd.ExecuteNonQuery();
                if (temp > 0)
                    va = true;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                sqlconn.Close();
            }
            return va;
        }



        /// <summary>
        /// 对Sql数据库执行事务处理，返回类型为bool
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool ExecutSqlTran(ArrayList list)
        {
            bool va = false;
            SqlConnection sqlconn = new SqlConnection(connSql);
            if (sqlconn.State != ConnectionState.Open)
                sqlconn.Open();
            SqlTransaction tran = sqlconn.BeginTransaction();
            SqlCommand cmd = new SqlCommand();
            try
            {
                cmd.Connection = sqlconn;
                cmd.Transaction = tran;
                for (int i = 0; i < list.Count; i++)
                {
                    cmd.CommandText = list[i].ToString();
                    cmd.ExecuteNonQuery();
                }
                tran.Commit();
                va = true;
            }
            catch (Exception e)
            {
                tran.Rollback();
                throw new Exception(e.Message);
            }
            finally
            {
                cmd.Dispose();
                sqlconn.Close();
            }
            return va;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool ExecutSqlTran(ArrayList list, ref string temp)
        {
            bool va = false;
            SqlConnection sqlconn = new SqlConnection(connSql);
            if (sqlconn.State != ConnectionState.Open)
                sqlconn.Open();
            SqlTransaction tran = sqlconn.BeginTransaction();
            SqlCommand cmd = new SqlCommand();
            try
            {
                cmd.Connection = sqlconn;
                cmd.Transaction = tran;
                for (int i = 0; i < list.Count; i++)
                {
                    cmd.CommandText = list[i].ToString();
                    temp = list[i].ToString();
                    cmd.ExecuteNonQuery();
                }
                tran.Commit();
                va = true;
            }
            catch
            {
                tran.Rollback();

            }
            finally
            {
                cmd.Dispose();
                sqlconn.Close();
            }
            return va;
        }

        /// <summary>
        /// 对Sql执行查询，返回第一行第一列的数据，返回类型为Object
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static object SelectOneRecord(string str)
        {
            object obj;
            SqlConnection sqlconn = new SqlConnection(connSql);
            SqlCommand cmd = new SqlCommand(str, sqlconn);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();
                obj = cmd.ExecuteScalar();
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                sqlconn.Close();
            }
            return obj;
        }

        /// <summary>
        /// 存储过程无参数查询，返回类型为DataSet
        /// </summary>
        /// <param name="strStore"></param>
        /// <returns></returns>
        public static DataSet ExecuteStoreForDs(string strStore)
        {
            SqlConnection sqlconn = new SqlConnection(connSql);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();

                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = strStore;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Connection = sqlconn;

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet("ds");
                da.Fill(ds);
                return ds;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                sqlconn.Close();
            }
        }

        /// <summary>
        /// 存储过程无参数查询，返回类型为SqlDataReader
        /// </summary>
        /// <param name="strStore"></param>
        /// <returns></returns>
        public static SqlDataReader ExecuteStoreForReader(string strStore)
        {
            SqlConnection sqlconn = new SqlConnection(connSql);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();

                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = strStore;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Connection = sqlconn;

                return cmd.ExecuteReader();
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                sqlconn.Close();
            }
        }

        /// <summary>
        /// 静态方法，将IDataReader转换成DataTable，返回类型为DataTable
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        public static DataTable ConvertDataReaderToDataTable(IDataReader reader)
        {
            DataTable objDataTable = new DataTable();
            int intFieldCount = reader.FieldCount;
            for (int intCounter = 0; intCounter < intFieldCount; ++intCounter)
            {
                objDataTable.Columns.Add(reader.GetName(intCounter), reader.GetFieldType(intCounter));
            }

            objDataTable.BeginLoadData();

            object[] objValues = new object[intFieldCount];
            while (reader.Read())
            {
                reader.GetValues(objValues);
                objDataTable.LoadDataRow(objValues, true);
            }
            reader.Close();
            objDataTable.EndLoadData();

            return objDataTable;
        }

        /// <summary>
        /// 静态方法，采用'MD5'算法，对输入密码进行加密
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public static String Encrypt(string password)
        {
            Byte[] clearBytes = new UnicodeEncoding().GetBytes(password);
            Byte[] hashedBytes = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(clearBytes);

            return BitConverter.ToString(hashedBytes).Substring(0, 14);
        }

        /// <summary>
        /// 递归查找控件的子控件，并清空内容
        /// </summary>
        /// <param name="contrs"></param>
        public static void SetControlEmpty(Control contrs)
        {
            foreach (Control con in contrs.Controls)
            {
                if (con is GridControl)
                    ((GridControl)con).DataSource = null;
                else if (con.HasChildren)
                {
                    SetControlEmpty(con);
                }
                else
                {
                    if (con.GetType().Name == "TextEdit")
                        con.Text = "";
                    //else if (con.GetType().Name == "ComboBox" && con.Visible != false)
                    //((ComboBox)con).SelectedIndex = -1;
                }
            }
        }
        /// <summary>
        /// 在DataGridView中查找值，然后选定当前值所在的行
        /// </summary>
        /// <param name="strKey"></param>
        /// <param name="dgv"></param>
        /// <param name="columnNo"></param>
        /// <returns></returns>
        public static int SelectStringInDGV(string strKey, DataGridView dgv, int columnNo)
        {
            int temp = 0;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv.Rows[i].Cells[columnNo].Value.ToString().Trim() == strKey)
                {
                    temp = i;
                    break;
                }
            }
            return temp;
        }


        public static int SelectStringInDGVTWO(string strKey1, string strKey2, DataGridView dgv, int columnNo1, int columnNo2)
        {
            int temp = 0;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv.Rows[i].Cells[columnNo1].Value.ToString() == strKey1 && dgv.Rows[i].Cells[columnNo2].Value.ToString() == strKey2)
                {
                    temp = i;
                    break;
                }
            }
            return temp;
        }

        /// <summary>
        /// 执行SqlCommand命令前的SqlCommand准备工作
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="cmdType"></param>
        /// <param name="cmdText"></param>
        /// <param name="para"></param>
        /// <param name="tran"></param>
        private static void PrepareCmd(SqlCommand cmd, CommandType cmdType, string cmdText, SqlParameter[] para, SqlTransaction tran)
        {
            SqlConnection sqlconn = new SqlConnection(connSql);
            cmd.Connection = sqlconn;
            cmd.CommandType = cmdType;
            cmd.CommandText = cmdText;
            if (tran != null)
            {
                cmd.Transaction = tran;
            }
            if (para != null && para.Length > 0)
            {
                foreach (SqlParameter pa in para)
                {
                    cmd.Parameters.Add(pa);
                }
            }
        }

        /// <summary>
        /// (通用)执行CMD.ExecuteNonQuery(),返回受影响的行数，返回类型为Int
        /// </summary>
        /// <param name="cmdType"></param>
        /// <param name="cmdText"></param>
        /// <param name="para"></param>
        /// <param name="tran"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(CommandType cmdType, string cmdText, SqlParameter[] para)
        {
            SqlCommand cmd = new SqlCommand();
            try
            {
                PrepareCmd(cmd, cmdType, cmdText, para, null);
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                int temp = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                return temp;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                cmd.Connection.Close();
            }
        }
        /// <summary>
        /// (通用) 通过SqlTransaction执行CMD.ExecuteNonQuery(),返回类型为bool;
        /// </summary>
        /// <param name="cmdType"></param>
        /// <param name="cmdText"></param>
        /// <param name="para"></param>
        /// <param name="tran"></param>
        /// <returns></returns>
        public static bool ExecuteNonQueryByTran(CommandType cmdType, string cmdText, SqlParameter[] para, SqlTransaction tran)
        {
            SqlCommand cmd = new SqlCommand();

            try
            {
                PrepareCmd(cmd, cmdType, cmdText, para, tran);
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                tran = cmd.Connection.BeginTransaction();
                cmd.ExecuteNonQuery();
                tran.Commit();
                return true;
            }
            catch
            {
                tran.Rollback();
                return false;
            }
            finally
            {
                cmd.Connection.Close();
            }
        }

        public static bool ExecuteNonQueryByTranNew(CommandType cmdType, string cmdText, SqlParameter[] para, SqlTransaction tran)
        {
            SqlCommand cmd = new SqlCommand();
            SqlConnection sqlconn = new SqlConnection(connSql);
            try
            {
                if (sqlconn.State != ConnectionState.Open)
                    sqlconn.Open();
                tran = sqlconn.BeginTransaction();

                cmd.Connection = sqlconn;
                cmd.CommandType = cmdType;
                cmd.CommandText = cmdText;
                if (tran != null)
                {
                    cmd.Transaction = tran;
                }
                if (para != null && para.Length > 0)
                {
                    foreach (SqlParameter pa in para)
                    {
                        cmd.Parameters.Add(pa);
                    }
                }

                cmd.ExecuteNonQuery();
                tran.Commit();
                return true;
            }
            catch (Exception ex)
            {
                tran.Rollback();
                return false;
            }
            finally
            {
                sqlconn.Close();
            }
        }

        /// <summary>
        /// （通用）执行CMD.ExecuteScalar返回第一条记录的第一列，返回类型为Object
        /// </summary>
        /// <param name="cmdType"></param>
        /// <param name="cmdText"></param>
        /// <param name="para"></param>
        /// <returns></returns>
        public static object ExecuteScalar(CommandType cmdType, string cmdText, SqlParameter[] para)
        {
            SqlCommand cmd = new SqlCommand();
            try
            {
                PrepareCmd(cmd, cmdType, cmdText, para, null);
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                object obj = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return obj;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                cmd.Connection.Close();
            }
        }
        /// <summary>
        /// (通用)执行CMD.ExecuteReader返回SqlDataReader
        /// </summary>
        /// <param name="cmdType"></param>
        /// <param name="cmdText"></param>
        /// <param name="para"></param>
        /// <returns></returns>
        public static SqlDataReader ExecuteReader(CommandType cmdType, string cmdText, SqlParameter[] para)
        {
            SqlCommand cmd = new SqlCommand();
            try
            {
                PrepareCmd(cmd, cmdType, cmdText, para, null);
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                SqlDataReader dataReader = cmd.ExecuteReader();
                cmd.Parameters.Clear();
                return dataReader;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                cmd.Connection.Close();
            }
        }

        /// <summary>
        /// (通用)执行SqlDataAdapter(CMD)返回Dataset
        /// </summary>
        /// <param name="cmdType"></param>
        /// <param name="cmdText"></param>
        /// <param name="para"></param>
        /// <returns></returns>
        public static DataSet DataAdapterByCmd(CommandType cmdType, string cmdText, SqlParameter[] para)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 60;
            try
            {
                PrepareCmd(cmd, cmdType, cmdText, para, null);
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                cmd.Parameters.Clear();
                return ds;
            }
            catch (SqlException e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                cmd.Connection.Close();
            }
        }
    }
}
