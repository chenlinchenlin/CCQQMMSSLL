using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data.OleDb;

namespace DX_QMS.Common
{
    class ExcelToSQL
    {
        public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
                //获取全部数据     
                string strConn = "Provider = Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";" + "Extended Properties = Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}$]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(ds, sheetName);

                //如果目标表不存在则创建,excel文件的第一行为列标题,从第二行开始全部都是数据记录     
                string strSql = string.Format("if not exists(select * from sysobjects where name = '{0}') create table {0}(", sheetName);   //以sheetName为表名     

                foreach (System.Data.DataColumn c in ds.Tables[0].Columns)
                {
                    strSql += string.Format("[{0}] varchar(255),", c.ColumnName);
                }
                strSql = strSql.Trim(',') + ")";

                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
                //用bcp导入数据        
                //excel文件中列的顺序必须和数据表的列顺序一致，因为数据导入时，是从excel文件的第二行数据开始，不管数据表的结构是什么样的，反正就是第一列的数据会插入到数据表的第一列字段中，第二列的数据插入到数据表的第二列字段中，以此类推，它本身不会去判断要插入的数据是对应数据表中哪一个字段的     

                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(connectionString))
                {
                    bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
                    bcp.BatchSize = 100;            //每次传输的行数        
                    bcp.NotifyAfter = 100;         //进度提示的行数        
                    bcp.DestinationTableName = sheetName;//目标表        
                    bcp.WriteToServer(ds.Tables[0]);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        //进度显示        
        void bcp_SqlRowsCopied(object sender, System.Data.SqlClient.SqlRowsCopiedEventArgs e)
        {
            //this.Text = e.RowsCopied.ToString();
            //this.Update();
        }






    }
}
