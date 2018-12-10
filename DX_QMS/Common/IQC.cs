using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace DX_QMS.Common
{
    class IQC
    {
        public int AddNewTestTypeRecord(string opertype, string testtype, string TTypes, string ifyiqi)
        {
            SqlParameter[] para = new SqlParameter[4];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@types", TTypes);
            para[3] = new SqlParameter("@Ifyiqi", ifyiqi);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestTypeRecord", para);
        }
        public DataSet SelectTestTypeRecord(string opertype, string testtype, string TTypes, string ifyiqi)
        {
            SqlParameter[] para = new SqlParameter[4];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@types", TTypes);
            para[3] = new SqlParameter("@Ifyiqi", ifyiqi);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestTypeRecord", para);

        }
        public int AddNewTestItemRecord(string opertype, string testtype, string testItem, int item, string oldtestitem)
        {
            SqlParameter[] para = new SqlParameter[5];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@item", item);
            para[4] = new SqlParameter("@oldtestitem", oldtestitem);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestItem", para);
        }
        public DataSet SelectTestItemRecord(string opertype, string testtype, string testItem, int item, string oldtestitem)
        {
            SqlParameter[] para = new SqlParameter[5];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@item", item);
            para[4] = new SqlParameter("@oldtestitem", oldtestitem);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestItem", para);
        }
        public int AddNewTestSubItemRecord(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue, string AQL, int item, int upscope, string oldtestsubitem)
        {
            SqlParameter[] para = new SqlParameter[15];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@AQL", AQL);
            para[12] = new SqlParameter("@item", item);
            para[13] = new SqlParameter("@upscope", upscope);
            para[14] = new SqlParameter("@oldtestsubitem", oldtestsubitem);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestSubItem", para);
        }
        public DataSet SelectTestSubItemRecord(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue, string AQL, int item, int upscope, string oldtestsubitem)
        {
            SqlParameter[] para = new SqlParameter[15];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@AQL", AQL);
            para[12] = new SqlParameter("@item", item);
            para[13] = new SqlParameter("@upscope", upscope);
            para[14] = new SqlParameter("@oldtestsubitem", oldtestsubitem);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestSubItem", para);
        }
        public int AddNewTestProgRecord(string opertype, string productcode, string checktype, string testtype, string PackType, string productname, string testItem, string TestSubItem, string TestDesc, string TestTool, string SampleType, string UpValue, string LowValue, float AQLValue, string AQL, int upscope, int testvalueqty, string unit, int checkcycle, int leadtime)
        {
            SqlParameter[] para = new SqlParameter[20];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@checktype", checktype);
            para[3] = new SqlParameter("@testtype", testtype);
            para[4] = new SqlParameter("@PackType", PackType);
            para[5] = new SqlParameter("@productname", productname);
            para[6] = new SqlParameter("@testItem", testItem);
            para[7] = new SqlParameter("@TestSubItem", TestSubItem);
            para[8] = new SqlParameter("@TestDesc", TestDesc);
            para[9] = new SqlParameter("@TestTool", TestTool);
            para[10] = new SqlParameter("@SampleType", SampleType);
            para[11] = new SqlParameter("@UpValue", UpValue);
            para[12] = new SqlParameter("@LowValue", LowValue);
            para[13] = new SqlParameter("@AQLValue", AQLValue);
            para[14] = new SqlParameter("@AQL", AQL);
            para[15] = new SqlParameter("@upscope", upscope);
            para[16] = new SqlParameter("@testvalueqty", testvalueqty);
            para[17] = new SqlParameter("@unit", unit);
            para[18] = new SqlParameter("@checkcycle", checkcycle);
            para[19] = new SqlParameter("@leadtime", leadtime);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestProg", para);
        }

        public int AddRepeatTestProg(string opertype, string productcode, string RepeatTestType, string testtype, string PackType, string productname, string testItem, string TestSubItem, string TestDesc, string TestTool, string SampleType, string UpValue, string LowValue, float AQLValue, string AQL, int upscope, int testvalueqty, string unit, int checkcycle, int leadtime)
        {
            SqlParameter[] para = new SqlParameter[20];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@RepeatTestType", RepeatTestType);
            para[3] = new SqlParameter("@testtype", testtype);
            para[4] = new SqlParameter("@PackType", PackType);
            para[5] = new SqlParameter("@productname", productname);
            para[6] = new SqlParameter("@testItem", testItem);
            para[7] = new SqlParameter("@TestSubItem", TestSubItem);
            para[8] = new SqlParameter("@TestDesc", TestDesc);
            para[9] = new SqlParameter("@TestTool", TestTool);
            para[10] = new SqlParameter("@SampleType", SampleType);
            para[11] = new SqlParameter("@UpValue", UpValue);
            para[12] = new SqlParameter("@LowValue", LowValue);
            para[13] = new SqlParameter("@AQLValue", AQLValue);
            para[14] = new SqlParameter("@AQL", AQL);
            para[15] = new SqlParameter("@upscope", upscope);
            para[16] = new SqlParameter("@testvalueqty", testvalueqty);
            para[17] = new SqlParameter("@unit", unit);
            para[18] = new SqlParameter("@checkcycle", checkcycle);
            para[19] = new SqlParameter("@leadtime", leadtime);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_RepeatTestProg", para);
        }


        public DataSet SelectTestProgRecord(string opertype, string productcode, string checktype, string testtype, string PackType, string productname, string testItem, string TestSubItem, string TestDesc, string TestTool, string SampleType, float UpValue, float LowValue, float AQLValue, string AQL, int upscope, int testvalueqty, string unit, int checkcycle, int leadtime)
        {
            SqlParameter[] para = new SqlParameter[20];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@checktype", checktype);
            para[3] = new SqlParameter("@testtype", testtype);
            para[4] = new SqlParameter("@PackType", PackType);
            para[5] = new SqlParameter("@productname", productname);
            para[6] = new SqlParameter("@testItem", testItem);
            para[7] = new SqlParameter("@TestSubItem", TestSubItem);
            para[8] = new SqlParameter("@TestDesc", TestDesc);
            para[9] = new SqlParameter("@TestTool", TestTool);
            para[10] = new SqlParameter("@SampleType", SampleType);
            para[11] = new SqlParameter("@UpValue", UpValue);
            para[12] = new SqlParameter("@LowValue", LowValue);
            para[13] = new SqlParameter("@AQLValue", AQLValue);
            para[14] = new SqlParameter("@AQL", AQL);
            para[15] = new SqlParameter("@upscope", upscope);
            para[16] = new SqlParameter("@testvalueqty", testvalueqty);
            para[17] = new SqlParameter("@unit", unit);
            para[18] = new SqlParameter("@checkcycle", checkcycle);
            para[19] = new SqlParameter("@leadtime", leadtime);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestProg", para);
        }

        public DataSet SelectRepeatTestProg(string opertype, string productcode, string RepeatTestType, string testtype, string PackType, string productname, string testItem, string TestSubItem, string TestDesc, string TestTool, string SampleType, float UpValue, float LowValue, float AQLValue, string AQL, int upscope, int testvalueqty, string unit, int checkcycle, int leadtime)
        {
            SqlParameter[] para = new SqlParameter[20];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@RepeatTestType", RepeatTestType);
            para[3] = new SqlParameter("@testtype", testtype);
            para[4] = new SqlParameter("@PackType", PackType);
            para[5] = new SqlParameter("@productname", productname);
            para[6] = new SqlParameter("@testItem", testItem);
            para[7] = new SqlParameter("@TestSubItem", TestSubItem);
            para[8] = new SqlParameter("@TestDesc", TestDesc);
            para[9] = new SqlParameter("@TestTool", TestTool);
            para[10] = new SqlParameter("@SampleType", SampleType);
            para[11] = new SqlParameter("@UpValue", UpValue);
            para[12] = new SqlParameter("@LowValue", LowValue);
            para[13] = new SqlParameter("@AQLValue", AQLValue);
            para[14] = new SqlParameter("@AQL", AQL);
            para[15] = new SqlParameter("@upscope", upscope);
            para[16] = new SqlParameter("@testvalueqty", testvalueqty);
            para[17] = new SqlParameter("@unit", unit);
            para[18] = new SqlParameter("@checkcycle", checkcycle);
            para[19] = new SqlParameter("@leadtime", leadtime);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_RepeatTestProg", para);
        }

        public int AddNewTestSamplePosition(string opertype, string productcode, string position, string validate, string paperposition, string sampletype, string sampledate,
                                            string qty, string developengineer, string sqeengineer, string remarks, string sup, string block, string userid, string ifhandle)
        {
            SqlParameter[] para = new SqlParameter[15];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@position", position);
            para[3] = new SqlParameter("@validate", validate);
            para[4] = new SqlParameter("@paperposition", paperposition);
            para[5] = new SqlParameter("@sampletype", sampletype);
            para[6] = new SqlParameter("@sampledate", sampledate);
            para[7] = new SqlParameter("@qty", qty);
            para[8] = new SqlParameter("@developEngineer", developengineer);
            para[9] = new SqlParameter("@SQEEngineer", sqeengineer);
            para[10] = new SqlParameter("@remarks", remarks);
            para[11] = new SqlParameter("@sup", sup);
            para[12] = new SqlParameter("@block", block);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@ifhandle", ifhandle);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestSamplePosition", para);
        }


        public int AddNewTestSamplePositionNew(string opertype, string productcode, string position, string validate, string paperposition, string sampletype,string sampleClass,string sampleCode, string sampledate,
                                    string qty, string developengineer, string sqeengineer, string remarks, string sup, string block, string userid, string ifhandle)
        {
            SqlParameter[] para = new SqlParameter[17];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@position", position);
            para[3] = new SqlParameter("@validate", validate);
            para[4] = new SqlParameter("@paperposition", paperposition);
            para[5] = new SqlParameter("@sampletype", sampletype);
            para[6] = new SqlParameter("@sampleClass", sampleClass);
            para[7] = new SqlParameter("@sampleCode", sampleCode);
            para[8] = new SqlParameter("@sampledate", sampledate);
            para[9] = new SqlParameter("@qty", qty);
            para[10] = new SqlParameter("@developEngineer", developengineer);
            para[11] = new SqlParameter("@SQEEngineer", sqeengineer);
            para[12] = new SqlParameter("@remarks", remarks);
            para[13] = new SqlParameter("@sup", sup);
            para[14] = new SqlParameter("@block", block);
            para[15] = new SqlParameter("@userid", userid);
            para[16] = new SqlParameter("@ifhandle", ifhandle);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestSamplePositionNew", para);
        }


        public DataSet SelectTestSamplePosition(string opertype, string productcode, string position, string validate, string paperposition, string sampletype, string sampledate,
                                           string qty, string developengineer, string sqeengineer, string remarks, string sup, string block, string userid, string ifhandle)
        {
            SqlParameter[] para = new SqlParameter[15];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@position", position);
            para[3] = new SqlParameter("@validate", validate);
            para[4] = new SqlParameter("@paperposition", paperposition);
            para[5] = new SqlParameter("@sampletype", sampletype);
            para[6] = new SqlParameter("@sampledate", sampledate);
            para[7] = new SqlParameter("@qty", qty);
            para[8] = new SqlParameter("@developEngineer", developengineer);
            para[9] = new SqlParameter("@SQEEngineer", sqeengineer);
            para[10] = new SqlParameter("@remarks", remarks);
            para[11] = new SqlParameter("@sup", sup);
            para[12] = new SqlParameter("@block", block);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@ifhandle", ifhandle);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestSamplePosition", para);
        }

        public DataSet SelectTestSamplePositionNew(string opertype, string productcode, string position, string validate, string paperposition, string sampletype,string sampleClass, string sampleCode, string sampledate,
                                   string qty, string developengineer, string sqeengineer, string remarks, string sup, string block, string userid, string ifhandle)
        {
            SqlParameter[] para = new SqlParameter[17];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@position", position);
            para[3] = new SqlParameter("@validate", validate);
            para[4] = new SqlParameter("@paperposition", paperposition);
            para[5] = new SqlParameter("@sampletype", sampletype);
            para[6] = new SqlParameter("@sampleClass", sampleClass);
            para[7] = new SqlParameter("@sampleCode", sampleCode);
            para[8] = new SqlParameter("@sampledate", sampledate);
            para[9] = new SqlParameter("@qty", qty);
            para[10] = new SqlParameter("@developEngineer", developengineer);
            para[11] = new SqlParameter("@SQEEngineer", sqeengineer);
            para[12] = new SqlParameter("@remarks", remarks);
            para[13] = new SqlParameter("@sup", sup);
            para[14] = new SqlParameter("@block", block);
            para[15] = new SqlParameter("@userid", userid);
            para[16] = new SqlParameter("@ifhandle", ifhandle);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestSamplePositionNew", para);
        }


        public string AddNewTestROHS(string opertype, string productcode, string guigei, string testdate, string sROHS, string YanWu, string userid, int item, string finaltestdate, string testresult, string testremarks, string filename)
        {
            SqlParameter[] para = new SqlParameter[13];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@guigei", guigei);
            para[3] = new SqlParameter("@testdate", testdate);
            para[4] = new SqlParameter("@sROHS", sROHS);
            para[5] = new SqlParameter("@YanWu", YanWu);
            para[6] = new SqlParameter("@userid", userid);
            para[7] = new SqlParameter("@item", item);
            para[8] = new SqlParameter("@finaltestdate", finaltestdate);
            para[9] = new SqlParameter("@testresult", testresult);
            para[10] = new SqlParameter("@testremarks", testremarks);
            para[11] = new SqlParameter("@filename", filename);
            para[12] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[12].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestROHS", para);
            return para[12].Value.ToString();
        }

        public DataSet SelectTestROHS(string opertype, string productcode, string guigei, string testdate, string sROHS, string YanWu, string userid, int item, string finaltestdate, string testresult, string testremarks, string filename)
        {
            SqlParameter[] para = new SqlParameter[13];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@guigei", guigei);
            para[3] = new SqlParameter("@testdate", testdate);
            para[4] = new SqlParameter("@sROHS", sROHS);
            para[5] = new SqlParameter("@YanWu", YanWu);
            para[6] = new SqlParameter("@userid", userid);
            para[7] = new SqlParameter("@item", item);
            para[8] = new SqlParameter("@finaltestdate", finaltestdate);
            para[9] = new SqlParameter("@testresult", testresult);
            para[10] = new SqlParameter("@testremarks", testremarks);
            para[11] = new SqlParameter("@filename", filename);
            para[12] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[12].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestROHS", para);
        }

        public string AddNewTestList(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue,
            string lotno, int qty, string userid, int sampleqty, int testqty, string productcode, string AQL, string TestResult, string TestFinalResult, string remarks, int item, string testvalue, string unit, string Checktype, string DCode, string Mdate, string ExpiryDate, string RSNO)
        {

            SqlParameter[] para = new SqlParameter[30];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@lotno", lotno);
            para[12] = new SqlParameter("@qty", qty);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@sampleqty", sampleqty);
            para[15] = new SqlParameter("@testqty", testqty);
            para[16] = new SqlParameter("@productcode", productcode);
            para[17] = new SqlParameter("@AQL", AQL);
            para[18] = new SqlParameter("@TestResult", TestResult);
            para[19] = new SqlParameter("@TestFinalResult", TestFinalResult);
            para[20] = new SqlParameter("@remarks", remarks);
            para[21] = new SqlParameter("@item", item);
            para[22] = new SqlParameter("@testvalue", testvalue);
            para[23] = new SqlParameter("@Unit", unit);
            para[24] = new SqlParameter("@CheckType", Checktype);
            para[25] = new SqlParameter("@DCode", DCode);
            para[26] = new SqlParameter("@Mdate", Mdate);
            para[27] = new SqlParameter("@ExpiryDate", ExpiryDate);
            para[28] = new SqlParameter("@rsno", RSNO);
            para[29] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[29].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestList", para);
            return para[29].Value.ToString();
        }


        public string AddNewTestListNew(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue,
    string lotno, int qty, string userid, int sampleqty, int testqty, string productcode, string AQL, string TestResult, string TestFinalResult, string remarks, int item, string testvalue, string unit, string Checktype, string DCode, string Mdate, string ExpiryDate, string RSNO,string NGtype,string Baddescribe,string manufacturer,int returnqty,string returnreason)
        {

            SqlParameter[] para = new SqlParameter[35];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@lotno", lotno);
            para[12] = new SqlParameter("@qty", qty);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@sampleqty", sampleqty);
            para[15] = new SqlParameter("@testqty", testqty);
            para[16] = new SqlParameter("@productcode", productcode);
            para[17] = new SqlParameter("@AQL", AQL);
            para[18] = new SqlParameter("@TestResult", TestResult);
            para[19] = new SqlParameter("@TestFinalResult", TestFinalResult);
            para[20] = new SqlParameter("@remarks", remarks);
            para[21] = new SqlParameter("@item", item);
            para[22] = new SqlParameter("@testvalue", testvalue);
            para[23] = new SqlParameter("@Unit", unit);
            para[24] = new SqlParameter("@CheckType", Checktype);
            para[25] = new SqlParameter("@DCode", DCode);
            para[26] = new SqlParameter("@Mdate", Mdate);
            para[27] = new SqlParameter("@ExpiryDate", ExpiryDate);
            para[28] = new SqlParameter("@rsno", RSNO);
            para[29] = new SqlParameter("@NGtype", NGtype);
            para[30] = new SqlParameter("@Baddescribe", Baddescribe);
            para[31] = new SqlParameter("@manufacturer", manufacturer);
            para[32] = new SqlParameter("@returnqty", returnqty);
            para[33] = new SqlParameter("@returnreason", returnreason);
            para[34] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[34].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertTestListNew", para);
            return para[34].Value.ToString();
        }


        public string AddReturnTestList(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue,
string lotno, int qty, string userid, int sampleqty, int testqty, string productcode, string AQL, string TestResult, string TestFinalResult, string remarks, int item, string testvalue, string unit, string Checktype, string DCode, string Mdate, string ExpiryDate, string RSNO, string NGtype, string Baddescribe, string manufacturer, int returnqty, string returnreason,string CheckItem ,string LotnoList)
        {

            SqlParameter[] para = new SqlParameter[37];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@lotno", lotno);
            para[12] = new SqlParameter("@qty", qty);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@sampleqty", sampleqty);
            para[15] = new SqlParameter("@testqty", testqty);
            para[16] = new SqlParameter("@productcode", productcode);
            para[17] = new SqlParameter("@AQL", AQL);
            para[18] = new SqlParameter("@TestResult", TestResult);
            para[19] = new SqlParameter("@TestFinalResult", TestFinalResult);
            para[20] = new SqlParameter("@remarks", remarks);
            para[21] = new SqlParameter("@item", item);
            para[22] = new SqlParameter("@testvalue", testvalue);
            para[23] = new SqlParameter("@Unit", unit);
            para[24] = new SqlParameter("@CheckType", Checktype);
            para[25] = new SqlParameter("@DCode", DCode);
            para[26] = new SqlParameter("@Mdate", Mdate);
            para[27] = new SqlParameter("@ExpiryDate", ExpiryDate);
            para[28] = new SqlParameter("@rsno", RSNO);
            para[29] = new SqlParameter("@NGtype", NGtype);
            para[30] = new SqlParameter("@Baddescribe", Baddescribe);
            para[31] = new SqlParameter("@manufacturer", manufacturer);
            para[32] = new SqlParameter("@returnqty", returnqty);
            para[33] = new SqlParameter("@returnreason", returnreason);
            para[34] = new SqlParameter("@CheckItem", CheckItem);
            para[35] = new SqlParameter("@LotnoList", LotnoList);
            para[36] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[36].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_ReturnTestList", para);
            return para[36].Value.ToString();
        }



        public DataSet SelectTestListNew(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue,
   string lotno, int qty, string userid, int sampleqty, int testqty, string productcode, string AQL, string TestResult, string TestFinalResult, string remarks, int item, string testvalue, string unit, string Checktype, string DCode, string Mdate, string ExpiryDate, string RSNO, string NGtype, string Baddescribe, string manufacturer, int returnqty, string returnreason)
        {

            SqlParameter[] para = new SqlParameter[35];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@lotno", lotno);
            para[12] = new SqlParameter("@qty", qty);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@sampleqty", sampleqty);
            para[15] = new SqlParameter("@testqty", testqty);
            para[16] = new SqlParameter("@productcode", productcode);
            para[17] = new SqlParameter("@AQL", AQL);
            para[18] = new SqlParameter("@TestResult", TestResult);
            para[19] = new SqlParameter("@TestFinalResult", TestFinalResult);
            para[20] = new SqlParameter("@remarks", remarks);
            para[21] = new SqlParameter("@item", item);
            para[22] = new SqlParameter("@testvalue", testvalue);
            para[23] = new SqlParameter("@Unit", unit);
            para[24] = new SqlParameter("@CheckType", Checktype);
            para[25] = new SqlParameter("@DCode", DCode);
            para[26] = new SqlParameter("@Mdate", Mdate);
            para[27] = new SqlParameter("@ExpiryDate", ExpiryDate);
            para[28] = new SqlParameter("@rsno", RSNO);
            para[29] = new SqlParameter("@NGtype", NGtype);
            para[30] = new SqlParameter("@Baddescribe", Baddescribe);
            para[31] = new SqlParameter("@manufacturer", manufacturer);
            para[32] = new SqlParameter("@returnqty", returnqty);
            para[33] = new SqlParameter("@returnreason", returnreason);
            para[34] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[34].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestListNew", para);
        }


        public DataSet SelectReturnTestList(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue,
string lotno, int qty, string userid, int sampleqty, int testqty, string productcode, string AQL, string TestResult, string TestFinalResult, string remarks, int item, string testvalue, string unit, string Checktype, string DCode, string Mdate, string ExpiryDate, string RSNO, string NGtype, string Baddescribe, string manufacturer, int returnqty, string returnreason, string CheckItem, string LotnoList)
        {
            SqlParameter[] para = new SqlParameter[37];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@lotno", lotno);
            para[12] = new SqlParameter("@qty", qty);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@sampleqty", sampleqty);
            para[15] = new SqlParameter("@testqty", testqty);
            para[16] = new SqlParameter("@productcode", productcode);
            para[17] = new SqlParameter("@AQL", AQL);
            para[18] = new SqlParameter("@TestResult", TestResult);
            para[19] = new SqlParameter("@TestFinalResult", TestFinalResult);
            para[20] = new SqlParameter("@remarks", remarks);
            para[21] = new SqlParameter("@item", item);
            para[22] = new SqlParameter("@testvalue", testvalue);
            para[23] = new SqlParameter("@Unit", unit);
            para[24] = new SqlParameter("@CheckType", Checktype);
            para[25] = new SqlParameter("@DCode", DCode);
            para[26] = new SqlParameter("@Mdate", Mdate);
            para[27] = new SqlParameter("@ExpiryDate", ExpiryDate);
            para[28] = new SqlParameter("@rsno", RSNO);
            para[29] = new SqlParameter("@NGtype", NGtype);
            para[30] = new SqlParameter("@Baddescribe", Baddescribe);
            para[31] = new SqlParameter("@manufacturer", manufacturer);
            para[32] = new SqlParameter("@returnqty", returnqty);
            para[33] = new SqlParameter("@returnreason", returnreason);
            para[34] = new SqlParameter("@CheckItem", CheckItem);
            para[35] = new SqlParameter("@LotnoList", LotnoList);
            para[36] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[36].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_ReturnTestList", para);
        }


        public DataSet SelectTestList(string opertype, string testtype, string testItem, string TestSubItem, string TestDesc, string TestTool, string PackType, string SampleType, float UpValue, float LowValue, float AQLValue,
           string lotno, int qty, string userid, int sampleqty, int testqty, string productcode, string AQL, string TestResult, string TestFinalResult, string remarks, int item, string testvalue, string unit, string Checktype, string DCode, string Mdate, string ExpiryDate, string RSNO)
        {

            SqlParameter[] para = new SqlParameter[30];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@TestSubItem", TestSubItem);
            para[4] = new SqlParameter("@TestDesc", TestDesc);
            para[5] = new SqlParameter("@TestTool", TestTool);
            para[6] = new SqlParameter("@PackType", PackType);
            para[7] = new SqlParameter("@SampleType", SampleType);
            para[8] = new SqlParameter("@UpValue", UpValue);
            para[9] = new SqlParameter("@LowValue", LowValue);
            para[10] = new SqlParameter("@AQLValue", AQLValue);
            para[11] = new SqlParameter("@lotno", lotno);
            para[12] = new SqlParameter("@qty", qty);
            para[13] = new SqlParameter("@userid", userid);
            para[14] = new SqlParameter("@sampleqty", sampleqty);
            para[15] = new SqlParameter("@testqty", testqty);
            para[16] = new SqlParameter("@productcode", productcode);
            para[17] = new SqlParameter("@AQL", AQL);
            para[18] = new SqlParameter("@TestResult", TestResult);
            para[19] = new SqlParameter("@TestFinalResult", TestFinalResult);
            para[20] = new SqlParameter("@remarks", remarks);
            para[21] = new SqlParameter("@item", item);
            para[22] = new SqlParameter("@testvalue", testvalue);
            para[23] = new SqlParameter("@Unit", unit);
            para[24] = new SqlParameter("@CheckType", Checktype);
            para[25] = new SqlParameter("@DCode", DCode);
            para[26] = new SqlParameter("@Mdate", Mdate);
            para[27] = new SqlParameter("@ExpiryDate", ExpiryDate);
            para[28] = new SqlParameter("@rsno", RSNO);
            para[29] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[29].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_InsertTestList", para);
        }

        public string Device_MaterialProdPic(string opertype, string PNo, string Pname, string userid, string sup, string checklist, string remarks)
        {

            SqlParameter[] para = new SqlParameter[8];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@pno", PNo);
            para[2] = new SqlParameter("@PName", Pname);
            para[3] = new SqlParameter("@userid", userid);
            para[4] = new SqlParameter("@sup", sup);
            para[5] = new SqlParameter("@checklist", checklist);
            para[6] = new SqlParameter("@remarks", remarks);
            para[7] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[7].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_ProdPictureAdd", para);
            return para[7].Value.ToString();
        }
        public DataSet SelectTestListRecord(string opertype, string testtype, string testItem, string TestSubItem, string TestTool, string lotno, string userid, string materialcode, string materialname, string vendorcode, string vendorname, string bdate, string edate)
        {

            SqlParameter[] para = new SqlParameter[13];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testitem", testItem);
            para[3] = new SqlParameter("@testsubitem", TestSubItem);
            para[4] = new SqlParameter("@testtools", TestTool);
            para[5] = new SqlParameter("@lotno", lotno);
            para[6] = new SqlParameter("@testuser", userid);
            para[7] = new SqlParameter("@materialcode", materialcode);
            para[8] = new SqlParameter("@materialname", materialname);
            para[9] = new SqlParameter("@vendorcode", vendorcode);
            para[10] = new SqlParameter("@vendorname", vendorname);
            para[11] = new SqlParameter("@testtime1", bdate);
            para[12] = new SqlParameter("@testtime2", edate);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_TestListSearchNew", para);
        }
        public void ExportToExcel(string fileFrom, string fileTo, DataTable dt, int colStar, int rowStar)
        {
            ExcelHelper excel = new ExcelHelper();
            try
            {
                excel.OpenExcelFile(fileFrom);
                excel.SaveAs(fileTo);
                excel.OpenExcelFile(fileTo);

                if (dt.Rows.Count < 1) return;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        excel.setCell(rowStar + i, colStar + j, dt.Rows[i][j].ToString());
                    }
                }
                excel.CurrentWorkBook.Save();
            }
            catch { }
            finally
            {
                excel.ReleaseExcel();
            }
        }

        public DataSet IfNorequireCheck(string testtype, string testItem, string TestSubItem, string productcode)
        {

            SqlParameter[] para = new SqlParameter[4];
            para[0] = new SqlParameter("@testtype", testtype);
            para[1] = new SqlParameter("@testitem", testItem);
            para[2] = new SqlParameter("@testsubitem", TestSubItem);
            para[3] = new SqlParameter("@productcode", productcode);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_TestRoHsLastTestDate", para);
        }

        public int IQC_ReliabilityOper(string opertype, string productcode, string Pname, string testtype, int testcycle, int leadtime, string userid)
        {

            SqlParameter[] para = new SqlParameter[7];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@productname", Pname);
            para[3] = new SqlParameter("@userid", userid);
            para[4] = new SqlParameter("@reliabilitytype", testtype);
            para[5] = new SqlParameter("@testcycle", testcycle);
            para[6] = new SqlParameter("@leadtime", leadtime);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_ReliabilitySetOper", para);
        }

        public string IQC_TestOtherInstock(string org_id, string item_id, string tran_id, string receptid, string productcode, string pqty, string enableqty, string userid, string lotno, string states, string remarks)
        {

            SqlParameter[] para = new SqlParameter[12];
            para[0] = new SqlParameter("@org_id", org_id);
            para[1] = new SqlParameter("@item_id", item_id);
            para[2] = new SqlParameter("@tran_id", tran_id);
            para[3] = new SqlParameter("@receptid", receptid);
            para[4] = new SqlParameter("@productcode", productcode);
            para[5] = new SqlParameter("@pqty", pqty);
            para[6] = new SqlParameter("@enableqty", enableqty);
            para[7] = new SqlParameter("@userid", userid);
            para[8] = new SqlParameter("@lotno", lotno);
            para[9] = new SqlParameter("@states", states);
            para[10] = new SqlParameter("@remarks", remarks);
            para[11] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[11].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Warehouse_IQCCheckInputAdd", para);
            return para[11].Value.ToString();
        }

        public int DailyCheck_ItemSetOper(string opertype, string section, string bigitem, string subitem, string oldsubitem, int sid, string dept)
        {
            SqlParameter[] para = new SqlParameter[7];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@section", section);
            para[2] = new SqlParameter("@bigitem", bigitem);
            para[3] = new SqlParameter("@subitem", subitem);
            para[4] = new SqlParameter("@oldsubitem", oldsubitem);
            para[5] = new SqlParameter("@sid", sid);
            para[6] = new SqlParameter("@dept", dept);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "DailyCheck_CheckItemOper", para);
        }
        public DataSet DailyCheck_StandardSetOper(string opertype, string section, string bigitem, string subitem, string pno, string standard)
        {
            SqlParameter[] para = new SqlParameter[6];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@section", section);
            para[2] = new SqlParameter("@bigitem", bigitem);
            para[3] = new SqlParameter("@subitem", subitem);
            para[4] = new SqlParameter("@pno", pno);
            para[5] = new SqlParameter("@standard", standard);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "DailyCheck_CheckStandardOper", para);
        }
        public int DailyCheck_StandardUpdateOper(string opertype, string section, string bigitem, string subitem, string pno, string standard)
        {
            SqlParameter[] para = new SqlParameter[6];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@section", section);
            para[2] = new SqlParameter("@bigitem", bigitem);
            para[3] = new SqlParameter("@subitem", subitem);
            para[4] = new SqlParameter("@pno", pno);
            para[5] = new SqlParameter("@standard", standard);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "DailyCheck_CheckStandardOper", para);
        }
        public DataSet DailyCheck_CheckResultOper(string opertype, string section, string org_id, string workno, string pno, string checkvalue, string checkresult, string bigitem, string subitem)
        {
            SqlParameter[] para = new SqlParameter[9];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@section", section);
            para[2] = new SqlParameter("@org_id", org_id);
            para[3] = new SqlParameter("@workno", workno);
            para[4] = new SqlParameter("@pno", pno);
            para[5] = new SqlParameter("@checkvalue", checkvalue);
            para[6] = new SqlParameter("@checkresult", checkresult);
            para[7] = new SqlParameter("@bigitem", bigitem);
            para[8] = new SqlParameter("@subitem", subitem);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "DailyCheck_CheckResultOper", para);
        }

        public int AddNewSMTBakeTimeRecord(string opertype, string productcode, string productname, decimal times, string userid)
        {
            SqlParameter[] para = new SqlParameter[5];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@productname", productname);
            para[3] = new SqlParameter("@times", times);
            para[4] = new SqlParameter("@userid", userid);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "SMT_PCBBakeTimeSetOper", para);
        }

        public int Add_QualitySoftVer(string opertype, string country, string customer, string machinetype, string fre, string icpcode, string softver, string userid)
        {
            SqlParameter[] para = new SqlParameter[8];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@country", country);
            para[2] = new SqlParameter("@customer", customer);
            para[3] = new SqlParameter("@machinetype", machinetype);
            para[4] = new SqlParameter("@fre", fre);
            para[5] = new SqlParameter("@icpcode", icpcode);
            para[6] = new SqlParameter("@softver", softver);
            para[7] = new SqlParameter("@userid", userid);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Quality_SoftVerOper", para);
        }
        public int AddNewESDItemRecord(string opertype, string testtype, string testItem, int item, string oldtestitem)
        {
            SqlParameter[] para = new SqlParameter[5];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@item", item);
            para[4] = new SqlParameter("@oldtestitem", oldtestitem);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "ESD_InsertESDItem", para);
        }
        public DataSet SelectESDItemRecord(string opertype, string testtype, string testItem, int item, string oldtestitem)
        {
            SqlParameter[] para = new SqlParameter[5];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@item", item);
            para[4] = new SqlParameter("@oldtestitem", oldtestitem);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "ESD_InsertESDItem", para);
        }
        public int AddNewESDProgRecord(string opertype, string block, string line, string testtype, string TestSubItem, string UpValue, string LowValue, int checkcycle, int leadtime, string userid)
        {
            SqlParameter[] para = new SqlParameter[10];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@block", block);
            para[2] = new SqlParameter("@line", line);
            para[3] = new SqlParameter("@testtype", testtype);
            para[4] = new SqlParameter("@TestSubItem", TestSubItem);
            para[5] = new SqlParameter("@UpValue", UpValue);
            para[6] = new SqlParameter("@LowValue", LowValue);
            para[7] = new SqlParameter("@checkcycle", checkcycle);
            para[8] = new SqlParameter("@leadtime", leadtime);
            para[9] = new SqlParameter("@userid", userid);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "ESD_InsertTestProg", para);
        }
        public string AddNewTestESD(string opertype, string block, string line, string testdate, string testtype, string testsubtype, string testvalue, string userid, int item, string finaltestdate, string testresult, string testremarks)
        {
            SqlParameter[] para = new SqlParameter[13];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@block", block);
            para[2] = new SqlParameter("@line", line);
            para[3] = new SqlParameter("@testdate", testdate);
            para[4] = new SqlParameter("@testtype", testtype);
            para[5] = new SqlParameter("@testsubtype", testsubtype);
            para[6] = new SqlParameter("@testvalue", testvalue);
            para[7] = new SqlParameter("@userid", userid);
            para[8] = new SqlParameter("@item", item);
            para[9] = new SqlParameter("@finaltestdate", finaltestdate);
            para[10] = new SqlParameter("@testresult", testresult);
            para[11] = new SqlParameter("@testremarks", testremarks);
            para[12] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[12].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "ESD_InsertTestESD", para);
            return para[12].Value.ToString();
        }
        public DataSet AddNewTestESDList(string opertype, string block, string line, string testdate, string testtype, string testsubtype, string testvalue, string userid, int item, string finaltestdate, string testresult, string testremarks)
        {
            SqlParameter[] para = new SqlParameter[13];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@block", block);
            para[2] = new SqlParameter("@line", line);
            para[3] = new SqlParameter("@testdate", testdate);
            para[4] = new SqlParameter("@testtype", testtype);
            para[5] = new SqlParameter("@testsubtype", testsubtype);
            para[6] = new SqlParameter("@testvalue", testvalue);
            para[7] = new SqlParameter("@userid", userid);
            para[8] = new SqlParameter("@item", item);
            para[9] = new SqlParameter("@finaltestdate", finaltestdate);
            para[10] = new SqlParameter("@testresult", testresult);
            para[11] = new SqlParameter("@testremarks", testremarks);
            para[12] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[12].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "ESD_InsertTestESD", para);
        }
        public int AddOQCCheckItem(string opertype, string testtype, string testItem, int item, string oldtestitem, string code)
        {
            SqlParameter[] para = new SqlParameter[6];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testItem", testItem);
            para[3] = new SqlParameter("@item", item);
            para[4] = new SqlParameter("@oldtestitem", oldtestitem);
            para[5] = new SqlParameter("@code", code);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "OQC_InsertCheckItem", para);
        }
        public string AddOQCTestRecord(string opertype, string productcode, string testitem, string cuscode, string customer, string org_id, string workno, string line,
              string testtool, string sendqty, string sampleqty, string AQLValue, string AQL, string testresult, string testremarks, string userid, string qc,
              string qe, string master, string states, string item, string latyper, string begindate, string enddate)
        {

            SqlParameter[] para = new SqlParameter[25];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@testitem", testitem);
            para[3] = new SqlParameter("@cuscode", cuscode);
            para[4] = new SqlParameter("@customer", customer);
            para[5] = new SqlParameter("@org_id", org_id);
            para[6] = new SqlParameter("@workno", workno);
            para[7] = new SqlParameter("@line", line);
            para[8] = new SqlParameter("@testtool", testtool);
            para[9] = new SqlParameter("@sendqty", sendqty);
            para[10] = new SqlParameter("@sampleqty", sampleqty);
            para[11] = new SqlParameter("@AQLValue", AQLValue);
            para[12] = new SqlParameter("@AQL", AQL);
            para[13] = new SqlParameter("@testresult", testresult);
            para[14] = new SqlParameter("@testremarks", testremarks);
            para[15] = new SqlParameter("@userid", userid);
            para[16] = new SqlParameter("@qc", qc);
            para[17] = new SqlParameter("@qe", qe);
            para[18] = new SqlParameter("@master", master);
            para[19] = new SqlParameter("@states", states);
            para[20] = new SqlParameter("@items", item);
            para[21] = new SqlParameter("@latyper", latyper);
            para[22] = new SqlParameter("@begindate", begindate);
            para[23] = new SqlParameter("@enddate", enddate);
            para[24] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[24].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "OQC_AddTestList", para);
            return para[24].Value.ToString();
        }
        public DataSet AddOQCTestRecordSearch(string opertype, string productcode, string testitem, string cuscode, string customer, string org_id, string workno, string line,
              string testtool, string sendqty, string sampleqty, string AQLValue, string AQL, string testresult, string testremarks, string userid, string qc,
              string qe, string master, string states, string item, string latyper, string begindate, string enddate)
        {

            SqlParameter[] para = new SqlParameter[25];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@testitem", testitem);
            para[3] = new SqlParameter("@cuscode", cuscode);
            para[4] = new SqlParameter("@customer", customer);
            para[5] = new SqlParameter("@org_id", org_id);
            para[6] = new SqlParameter("@workno", workno);
            para[7] = new SqlParameter("@line", line);
            para[8] = new SqlParameter("@testtool", testtool);
            para[9] = new SqlParameter("@sendqty", sendqty);
            para[10] = new SqlParameter("@sampleqty", sampleqty);
            para[11] = new SqlParameter("@AQLValue", AQLValue);
            para[12] = new SqlParameter("@AQL", AQL);
            para[13] = new SqlParameter("@testresult", testresult);
            para[14] = new SqlParameter("@testremarks", testremarks);
            para[15] = new SqlParameter("@userid", userid);
            para[16] = new SqlParameter("@qc", qc);
            para[17] = new SqlParameter("@qe", qe);
            para[18] = new SqlParameter("@master", master);
            para[19] = new SqlParameter("@states", states);
            para[20] = new SqlParameter("@items", item);
            para[21] = new SqlParameter("@latyper", latyper);
            para[22] = new SqlParameter("@begindate", begindate);
            para[23] = new SqlParameter("@enddate", enddate);
            para[24] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[24].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "OQC_AddTestList", para);
        }

        public string QMS_MenuList(string opertype, string module,string moduleGroup, string menuname,string formname, string formhead,string formdescribe)
        {
            SqlParameter[] para = new SqlParameter[8];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@module", module);
            para[2] = new SqlParameter("@moduleGroup", moduleGroup);
            para[3] = new SqlParameter("@menuname", menuname);
            para[4] = new SqlParameter("@formname", formname);
            para[5] = new SqlParameter("@formhead", formhead);
            para[6] = new SqlParameter("@formdescribe", formdescribe);
            para[7] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[7].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "QMS_MenuList", para);
            return para[7].Value.ToString();
        }




    }
}
