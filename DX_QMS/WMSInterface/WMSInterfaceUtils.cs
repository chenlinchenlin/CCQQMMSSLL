using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMSInterface;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Data;
using Newtonsoft.Json.Linq;
using System.Collections;
using System.Web.Script.Serialization;
using DX_QMS.Common;

namespace barcode.WMSInterface
{
    class WMSInterfaceUtils
    {
        //通用接口方法
        public static bool CallErpInterface(string batchNumber, string method, string requestData, ref string responseData)
        {
            bool restult = false;
            try
            {
                WIInput myInput = new WIInput();
                myInput.P_BATCH_NUMBER = batchNumber;
                myInput.P_IFACE_CODE = method;
                myInput.P_REQUEST_DATA = requestData;
                WIHeader myHeader = new WIHeader();
                myHeader.xmlns = "http://xmlns.oracle.com/apps/per/rest/comm_ws/header";
                myHeader.Responsibility = "CUX_DEVE_RESP";
                myHeader.RespApplication = "CUX";
                myHeader.SecurityGroup = "STANDARD";
                myHeader.NLSLanguage = "AMERICAN";
                myHeader.Org_Id = "0";

                WIRequst myRequst = new WIRequst();
                myRequst.InputParameters = myInput;
                myRequst.RESTHeader = myHeader;
                myRequst.@xmlns = "http://xmlns.oracle.com/apps/per/rest/comm_ws/invokeebsws";
                WsRequest wsRequest = new WsRequest();
                wsRequest.comm_ws = myRequst;
                string request = JsonConvert.SerializeObject(wsRequest);
                request = request.Insert(request.IndexOf("xmlns"), "@");
                HttpClient client = new HttpClient();
                client.Timeout = TimeSpan.FromMilliseconds(10000000);
                //Uri uri = new Uri("http://192.168.2.139:8004/webservices/rest/comm_ws/invokeebsws/");
                Uri uri = new Uri("http://192.168.2.120:8000/webservices/rest/comm_ws/invokeebsws/");
                //Uri uri = new Uri("http://192.168.2.139:8010/webservices/rest/comm_ws/invokeebsws/");
                client.BaseAddress = uri;
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes("REST:Hytera1993")));
                var content = new StringContent(request, Encoding.UTF8, "application/json");
                HttpResponseMessage response = client.PostAsync(uri, content).Result;
                string responseJsonText = response.Content.ReadAsStringAsync().Result;
                int successIndex = responseJsonText.IndexOf("\"X_RETURN_CODE\" : \"S1001000\"");
                if (0 < successIndex)
                {
                    restult = true;
                }
                int responseEnd = responseJsonText.LastIndexOf("\"");
                int responseStart = responseJsonText.IndexOf("X_RESPONSE_DATA");
                //byte[] bytes = Convert.FromBase64String(responseJsonText.Substring(responseStart + 20, responseEnd - responseStart - 20));
                responseData = Encoding.UTF8.GetString(Convert.FromBase64String(responseJsonText.Substring(responseStart + 20, responseEnd - responseStart - 20)));
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
            return restult;

        }




        public static bool CallWMSInterface(string batchNumber, string method, string requestData, ref string responseData,ref string responseoriginalDate)
        {
            bool restult = false;
            try
            {
                WIInput myInput = new WIInput();
                myInput.P_BATCH_NUMBER = batchNumber;
                myInput.P_IFACE_CODE = method;
                myInput.P_REQUEST_DATA = requestData;
                WIHeader myHeader = new WIHeader();
                myHeader.xmlns = "http://xmlns.oracle.com/apps/per/rest/comm_ws/header";
                myHeader.Responsibility = "CUX_DEVE_RESP";
                myHeader.RespApplication = "CUX";
                myHeader.SecurityGroup = "STANDARD";
                myHeader.NLSLanguage = "AMERICAN";
                myHeader.Org_Id = "0";

                WIRequst myRequst = new WIRequst();
                myRequst.InputParameters = myInput;
                myRequst.RESTHeader = myHeader;
                myRequst.@xmlns = "http://xmlns.oracle.com/apps/per/rest/comm_ws/invokeebsws";
                WsRequest wsRequest = new WsRequest();
                wsRequest.comm_ws = myRequst;
                string request = JsonConvert.SerializeObject(wsRequest);
                request = request.Insert(request.IndexOf("xmlns"), "@");
                HttpClient client = new HttpClient();
                client.Timeout = TimeSpan.FromMilliseconds(10000000);
                //Uri uri = new Uri("http://192.168.2.139:8004/webservices/rest/comm_ws/invokeebsws/");
                Uri uri = new Uri("http://192.168.2.120:8000/webservices/rest/comm_ws/invokeebsws/");
                //Uri uri = new Uri("http://192.168.2.139:8010/webservices/rest/comm_ws/invokeebsws/");
                client.BaseAddress = uri;
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes("REST:Hytera1993")));
                var content = new StringContent(request, Encoding.UTF8, "application/json");
                HttpResponseMessage response = client.PostAsync(uri, content).Result;
                string responseJsonText = response.Content.ReadAsStringAsync().Result;

                responseoriginalDate = responseJsonText;

                int successIndex = responseJsonText.IndexOf("\"X_RETURN_CODE\" : \"S1001000\"");
                if (0 < successIndex)
                {
                    restult = true;
                }
                int responseEnd = responseJsonText.LastIndexOf("\"");
                int responseStart = responseJsonText.IndexOf("X_RESPONSE_DATA");
                //byte[] bytes = Convert.FromBase64String(responseJsonText.Substring(responseStart + 20, responseEnd - responseStart - 20));
                responseData = Encoding.UTF8.GetString(Convert.FromBase64String(responseJsonText.Substring(responseStart + 20, responseEnd - responseStart - 20)));
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
            return restult;

        }






        public static bool CallErpInterface2(string batchNumber, string method, string requestData, ref string responseData)
        {
            bool restult = false;
            WMSWorkPick oo=null;
            try
            {
                WIInput myInput = new WIInput();
                myInput.P_BATCH_NUMBER = batchNumber;
                myInput.P_IFACE_CODE = method;
                myInput.P_REQUEST_DATA = requestData;
                WIHeader myHeader = new WIHeader();
                myHeader.xmlns = "http://xmlns.oracle.com/apps/per/rest/comm_ws/header";
                myHeader.Responsibility = "CUX_DEVE_RESP";
                myHeader.RespApplication = "CUX";
                myHeader.SecurityGroup = "STANDARD";
                myHeader.NLSLanguage = "AMERICAN";
                myHeader.Org_Id = "0";

                WIRequst myRequst = new WIRequst();
                myRequst.InputParameters = myInput;
                myRequst.RESTHeader = myHeader;
                myRequst.@xmlns = "http://xmlns.oracle.com/apps/per/rest/comm_ws/invokeebsws";
                WsRequest wsRequest = new WsRequest();
                wsRequest.comm_ws = myRequst;
                string request = JsonConvert.SerializeObject(wsRequest);
                request = request.Insert(request.IndexOf("xmlns"), "@");
                HttpClient client = new HttpClient();
                //Uri uri = new Uri("http://192.168.2.139:8004/webservices/rest/comm_ws/invokeebsws/");
                Uri uri = new Uri("http://192.168.2.120:8000/webservices/rest/comm_ws/invokeebsws/");
                //Uri uri = new Uri("http://192.168.2.139:8010/webservices/rest/comm_ws/invokeebsws/");
                client.BaseAddress = uri;
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes("REST:Hytera1993")));
                var content = new StringContent(request, Encoding.UTF8, "application/json");
                HttpResponseMessage response = client.PostAsync(uri, content).Result;
                string responseJsonText = response.Content.ReadAsStringAsync().Result;
                int successIndex = responseJsonText.IndexOf("\"X_RETURN_CODE\" : \"S1001000\"");
                if (0 < successIndex)
                {
                    restult = true;
                }
                int responseEnd = responseJsonText.LastIndexOf("\"");
                int responseStart = responseJsonText.IndexOf("X_RESPONSE_DATA");
                responseData = Encoding.UTF8.GetString(Convert.FromBase64String(responseJsonText.Substring(responseStart + 20, responseEnd - responseStart - 20)));
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
            return true;

        }
        public static DataTable GetWorkLotInfo(string dept,string orgid,string workno)
        {
            string sql="";
            if (dept == "Assembly")
            {
                //sql = "select flotno lotno,a.qty,frerange,machinetype from Assembly_Serialbuild a where a.org_id='" + orgid + "' and a.workno='" + workno + "' ";
                //sql += " group by flotno,a.qty,frerange,machinetype ";
                sql=" select lotno,count(lotno) qty,frerange,machinetype from Assembly_Prod a inner join Ass_Maintain ";
                sql+=" d on a.org_id=d.org_id and a.workno=d.workno ";
                sql += " where a.org_id='" + orgid + "' and a.workno='" + workno + "' group by lotno,a.qty,frerange,machinetype ";
            }
            else if (dept == "Pack")
            {
                sql = "if exists(select 1 from Pack_Maintain where org_id='" + orgid + "' and workno='" + workno + "' and (serialno like 'Z%' or serialno like 'R%') )";
                sql += "select flotno lotno,a.qty,frerange,machinetype from Pack_Serialbuild a where a.org_id='" + orgid + "' and a.workno='" + workno + "' and (firstno like 'Z%' or firstno like'R%') ";
                sql += " group by flotno,a.qty,frerange,machinetype  else if exists(select 1 from Pack_Serialbuild where org_id='"+orgid+"' and workno='"+workno+"')";
                sql += "select flotno lotno,a.qty,frerange,machinetype from Pack_Serialbuild a where a.org_id='" + orgid + "' and a.workno='" + workno + "'";
                sql += " group by flotno,a.qty,frerange,machinetype  else";
                sql += " select lotno,count(lotno) qty,'' frerange,'' machinetype from Pack_maintain where org_id='"+orgid+"' and workno='"+workno+"' group by lotno";
            }
            else if (dept == "SYS")
            {
                //sql = "select lotno,p.qty,'' frerange,'' machinetype from SYS_Maintain a inner join SYS_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno where a.org_id='" + orgid + "' and a.workno='" + workno + "' ";
                //sql += " group by lotno,p.qty ";
                sql = "select lotno,count(lotno) qty,'' frerange,'' machinetype from SYS_Maintain a where a.org_id='" + orgid + "' and a.workno='" + workno + "' ";
                sql += " group by lotno";
            }
            else if (dept == "SMT")
            {
                //sql = "select lotno,p.qty,'' frerange,'' machinetype from SMT_Maintain a inner join SMT_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno where a.org_id='" + orgid + "' and a.workno='" + workno + "' ";
                //sql += " group by lotno,p.qty,frerange,machinetype ";

                sql=" select lotno,count(lotno) qty,'' frerange,'' machinetype from SMT_Maintain a ";
                sql += " where a.org_id='" + orgid + "' and a.workno='" + workno + "' group by lotno ";
            }
            else if (dept == "SPLIT")
            {
                sql = " select lotno,count(lotno) qty,'' frerange,'' machinetype,workno from SMT_Maintain a ";
                sql += " where a.org_id='" + orgid + "' and a.lotno='" + workno + "' group by lotno,workno ";
            }

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        public static DataTable GetWorkLotInfoByLot(string dept, string orgid, string lotno)
        {
            string sql = "";
            if (dept == "Assembly")
            {
                sql = "if not exists(select 1 from lotnoinfo where lotno='"+lotno+"' )select a.workno,lotno,p.qty,frerange,machinetype from Ass_Maintain a inner join Assembly_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno where a.org_id='" + orgid + "' and a.lotno='" + lotno + "' ";
                sql += " group by a.workno,lotno,p.qty,frerange,machinetype ";
                sql += " else ";
                sql += "select a.workno,a.lotno,l.qty,frerange,machinetype from Ass_Maintain a inner join Assembly_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno inner join lotnoinfo l on a.org_id=l.org_id and a.lotno=l.lotno where a.org_id='" + orgid + "' and a.lotno='" + lotno + "' ";
                sql += " group by a.workno,a.lotno,l.qty,frerange,machinetype ";
            }
            else if (dept == "SMT")
            {
                sql = "if not exists(select 1 from lotnoinfo where lotno='" + lotno + "' )select a.workno,lotno,p.qty,frerange,machinetype from SMT_Maintain a inner join SMT_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno where a.org_id='" + orgid + "' and a.lotno='" + lotno + "' ";
                sql += " group by a.workno,lotno,p.qty,frerange,machinetype ";
                sql += " else ";
                sql += "select a.workno,a.lotno,l.qty,frerange,machinetype from SMT_Maintain a inner join SMT_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno inner join lotnoinfo l on a.org_id=l.org_id and a.lotno=l.lotno where a.org_id='" + orgid + "' and a.lotno='" + lotno + "' ";
                sql += " group by a.workno,a.lotno,l.qty,frerange,machinetype ";
            }
            else if (dept == "Pack")
            {
                sql = "if exists(select 1 from Pack_Maintain where org_id='" + orgid + "' and lotno='" + lotno + "' and (left(serialno,1)='Z' or left(serialno,1)='R') )";
                sql = "select a.workno,flotno lotno,qty,frerange,machinetype from Pack_Serialbuild a where a.org_id='" + orgid + "' and a.flotno='" + lotno + "' and (left(firstno,1)='Z' or left(firstno,1)='R') ";
                sql += " group by a.workno,flotno,qty,frerange,machinetype  else ";
                sql = "select a.workno,flotno lotno,qty,frerange,machinetype from Pack_Serialbuild a where a.org_id='" + orgid + "' and a.flotno='" + lotno + "'";
                sql += " group by a.workno,flotno,qty,frerange,machinetype ";
            }
            else if (dept == "SYS")
            {
                sql = "if not exists(select 1 from lotnoinfo where lotno='" + lotno + "' )select a.workno,lotno,p.qty,'' frerange,'' machinetype from SYS_Maintain a inner join SYS_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno where a.org_id='" + orgid + "' and a.lotno='" + lotno + "' ";
                sql += " group by a.workno,lotno,p.qty ";
                sql += " else ";
                sql += "select a.workno,a.lotno,l.qty,frerange,machinetype from SYS_Maintain a inner join SYS_Prod p on a.org_id=p.org_id and a.workno=p.workno inner join lotnoinfo l on a.org_id=l.org_id and a.lotno=l.lotno where a.org_id='" + orgid + "' and a.lotno='" + lotno + "' ";
                sql += " group by a.workno,a.lotno,l.qty,frerange,machinetype ";
            }
            else if (dept == "TEST")
            {
                sql = "select max(a.workno) workno,BoxNo lotno,count(SN) qty,max(p.frerange) frerange,max(p.machinetype) machinetype from SMT_TestPackCheck a left join SMT_Prod p on a.org_id=p.org_id and a.workno=p.workno where BoxNo='" + lotno + "' group by BoxNo";
            }
            else if (dept == "SYSTEM")
            {
                //sql = "select TOP 1 isnull(WO,workno) workno,subserialno lotno,ISNULL(MO,mono) MO,qty,ISNULL(Module,PraecommModel) module,Package,SetNo,remarks,packageno from PackingList where subserialno='" + lotno + "'";
                sql = "select TOP 1 workno,subserialno lotno,mono MO,qty,Praecommtype module,Package,SetNo,remarks,packageno from PackingList where subserialno='" + lotno + "'";
            }
            else if (dept == "Cancel")
            {
                //sql = "select TOP 1 isnull(WO,workno) workno,subserialno lotno,ISNULL(MO,mono) MO,qty,ISNULL(Module,PraecommModel) module,Package,SetNo,remarks,packageno from PackingListHistory where subserialno='" + lotno + "'";
                sql = "select TOP 1 workno,subserialno lotno,mono MO,qty,Praecommtype module,Package,SetNo,remarks,packageno from PackingListHistory where subserialno='" + lotno + "'";
            }
            else if (dept == "Rework")
            {
                sql += "select a.workno,a.lotno,a.qty,'' frerange,'' machinetype from lotnoinfo a where a.org_id='" + orgid + "' and a.lotno='" + lotno + "' ";
                sql += " group by a.workno,a.lotno,a.qty ";
            }
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        public static DataTable GetSNInfoByLot(string dept, string lotno)
        {
            string sql = "";
            if (dept == "Assembly")
            {
                sql = "select serialno,p.productcode,p.prodes from Ass_Maintain a inner join Assembly_Prod p on a.org_id=p.org_id and a.workno=p.workno where lotno='" + lotno + "' ";
            }
            else if (dept == "SMT")
            {
                sql = "select serialno,p.productcode,p.prodes from SMT_Maintain a inner join SMT_Prod p on a.org_id=p.org_id and a.workno=p.workno where lotno='" + lotno + "' ";
            }
            else if (dept == "SYS")
            {
                sql = "select serialno,p.productcode,p.prodes from SYS_Maintain a inner join SYS_Prod p on a.org_id=p.org_id and a.workno=p.workno where lotno='" + lotno + "' ";
            }
            else if (dept == "Pack")
            {
                sql = "if exists(select 1 from Pack_Maintain where lotno='" + lotno + "' and (left(serialno,1)='Z' or left(serialno,1)='R') )";
                sql += "select serialno,p.productcode,p.prodes from Pack_Maintain a inner join Pack_Prod p on a.org_id=p.org_id and a.workno=p.workno and p.countrytype in('EBhost','Discount','glocalhost','GESN','ABCard','ABHand') where lotno='" + lotno + "' and (left(serialno,1)='Z' or left(serialno,1)='R') ";
                sql += " else if exists(select serialno,p.productcode,m.materialname prodes from Pack_Maintain a inner join Pack_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno and a.lotno=p.flotno left join MaterialSpec m on p.productcode=m.materialcode where lotno='" + lotno + "') ";
                sql += " select serialno,p.productcode,m.materialname prodes from Pack_Maintain a inner join Pack_Serialbuild p on a.org_id=p.org_id and a.workno=p.workno and a.lotno=p.flotno left join MaterialSpec m on p.productcode=m.materialcode where lotno='" + lotno + "' ";
                sql += " else if exists(select 1 from Pack_Serialbuild where flotno='" + lotno + "') and not exists(select 1 from Pack_Maintain where lotno='"+lotno+"') ";
                sql += " select '' serialno,productcode,'' prodes from Pack_Serialbuild where flotno='" + lotno + "' ";
                sql += " else begin declare @p varchar(50) select @p=productcode from Pack_Prod where workno in(select workno from Pack_Maintain where lotno='"+lotno+"') select serialno,@p productcode,'' prodes from Pack_Maintain where lotno='" + lotno + "' end ";
            }
            else if (dept == "SYSTEM")
            {
                //sql = "select code,PraecommDescription,SetNo,remarks,Module,Package from PackingList where subserialno='" + lotno + "'";PraecommDescription prodes
                sql = "select workno,subserialno lotno,mono MO,qty,Praecommtype Module,Package,SetNo,remarks,packageno,code productcode,PraecommDescription prodes from PackingList where subserialno='" + lotno + "'";
            }

            else if (dept == "Cancel")
            {
                //sql = "select code,PraecommDescription,SetNo,remarks,Module,Package from PackingListHistory where subserialno='" + lotno + "'";
                sql = "select workno,subserialno lotno,mono MO,qty,Praecommtype Module,Package,SetNo,remarks,packageno,code productcode,PraecommDescription prodes from PackingListHistory where subserialno='" + lotno + "'";

            }
            else if (dept == "TEST")
            {
                sql = "select SN serialno,p.productcode,p.prodes from SMT_TestPackCheck a left join SMT_Prod p on a.org_id=p.org_id and a.workno=p.workno where BoxNo='" + lotno + "' ";
            }
            else if (dept == "Rework")
            {
                sql = "select serialno,p.productcode,'' prodes from SAP_LotRelationserialnolog a inner join lotnoinfo p on a.lotno=p.lotno where a.lotno='" + lotno + "' ";
            }
            else if (dept == "OEM")
            {
                sql = "select workno,lotno,productname prodes,productcode,qty,machinetype,remarks,orderno mo,org_id from OEM_EMSLotNo where lotno='" + lotno + "' ";
            }
            else if (dept == "SPLIT")
            {
                sql = "select serialno,p.productcode,p.prodes from SMT_Maintain a inner join SMT_Prod p on a.org_id=p.org_id and a.workno=p.workno where lotno='" + lotno + "' ";
            }
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        public static void LotNoWriteERP(string dept,string orgid,string workno)
        {
            WMSLPNSnItem wmslpnSnItem = new WMSLPNSnItem();
            List<WMSLPNSnItem> list = new List<WMSLPNSnItem>();
            WMSLPNItem wmslpnItem = new WMSLPNItem();
            List<WMSLPNItem> wmsLPNItemList = new List<WMSLPNItem>();

            DataTable DtLot = WMSInterface.WMSInterfaceUtils.GetWorkLotInfo(dept, orgid, workno);
            for (int i = 0; i < DtLot.Rows.Count; i++)
            {
                DataTable DtSn = WMSInterface.WMSInterfaceUtils.GetSNInfoByLot("Assembly", DtLot.Rows[i]["lotno"].ToString());
                for (int j = 0; j < DtSn.Rows.Count; j++)
                {
                    //生成每个产品的信息
                    wmslpnSnItem.itemSize = "1";
                    wmslpnSnItem.itemDesc = DtSn.Rows[j]["prodes"].ToString();
                    wmslpnSnItem.itemNumber = DtSn.Rows[j]["productcode"].ToString();
                    wmslpnSnItem.SN = DtSn.Rows[j]["serialno"].ToString();
                    //将两个产品加入List
                    list.Add(wmslpnSnItem);
                }
                //生成批次（装配、包装的箱号批次，系统车间的装箱批次）

                wmslpnItem.LPN = DtLot.Rows[i]["lotno"].ToString(); //大箱批次号
                wmslpnItem.LPNSize = DtLot.Rows[i]["qty"].ToString(); //大箱容量
                wmslpnItem.setNo = "";  //系统车间生成的才填
                wmslpnItem.remark = DtLot.Rows[i]["frerange"].ToString();
                wmslpnItem.module = DtLot.Rows[i]["machinetype"].ToString();
                wmslpnItem.packingName = ""; //包装规格
                wmslpnItem.SNItems = list;


                wmsLPNItemList.Add(wmslpnItem);

            }
            //生成头文件，包含工单信息
            WMSLPNHeader wmslpnHeader = new WMSLPNHeader();
            wmslpnHeader.action = "CREATE";   // "CREATE"/"DELETE"
            wmslpnHeader.LPNQty = DtLot.Rows.Count.ToString(); //批次数量，有几个大箱就是几
            wmslpnHeader.organizationCode = orgid;//组织代号
            if (dept != "SYS")
            {
                wmslpnHeader.sourceCode = "OTHERS";// "OTHERS"/"SYSTEM"     系统的装箱清单用SYSTEM，其他车间用OTHERS
                wmslpnHeader.moNumber = "";
            }
            else
            {
                wmslpnHeader.sourceCode = "SYSTEM";
                wmslpnHeader.moNumber = "";
            }
            wmslpnHeader.LPNItems = wmsLPNItemList;
            wmslpnHeader.workOrder = workno;//工单号

            WMSLPNInputParamData wmslpnInputParamData = new WMSLPNInputParamData();
            wmslpnInputParamData.header = wmslpnHeader;

            WMSLPNRequestData wmsLPNRequestData = new WMSLPNRequestData();
            wmsLPNRequestData.wsinterface = wmslpnInputParamData;

            string batchNumn = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string method = "HYT_LPN_WS";
            string requestData = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(wmsLPNRequestData)));
            string responseData = string.Empty;
            bool result = WMSInterface.WMSInterfaceUtils.CallErpInterface(batchNumn, method, requestData, ref responseData);
            string s = "0";
            if (result)
                s = "1";
            else
                s = "0";
            string sqlupdate = " update lotnoinfo set IfSucess='" + s + "' where workno='"+workno+"' and org_id='"+orgid+"'";
            sqlupdate += " insert into WMSInterfaceLog(org_id, workno, batchNumn, method, requestData, result,opertype, uploaddate)";
            sqlupdate += "values('"+orgid+"','"+workno+"','" + batchNumn + "','" + method + "','" + responseData + "','" + s + "','" + wmslpnHeader.action + "',getdate())";
            DbAccess.ExecuteSql(sqlupdate);
        }
        public static void LotNoWriteERPByLot(string dept,string orgid,string lotno)
        {
            WMSLPNSnItem wmslpnSnItem = new WMSLPNSnItem();
            List<WMSLPNSnItem> list = new List<WMSLPNSnItem>();
            WMSLPNItem wmslpnItem = new WMSLPNItem();
            List<WMSLPNItem> wmsLPNItemList = new List<WMSLPNItem>();

            DataTable DtLot = WMSInterface.WMSInterfaceUtils.GetWorkLotInfoByLot(dept,orgid,lotno);
            for (int i = 0; i < DtLot.Rows.Count; i++)
            {
                DataTable DtSn = WMSInterface.WMSInterfaceUtils.GetSNInfoByLot("Assembly", DtLot.Rows[i]["lotno"].ToString());
                for (int j = 0; j < DtSn.Rows.Count; j++)
                {
                    //生成每个产品的信息
                    wmslpnSnItem.itemSize = "1";
                    wmslpnSnItem.itemDesc = DtSn.Rows[j]["prodes"].ToString();
                    wmslpnSnItem.itemNumber = DtSn.Rows[j]["productcode"].ToString();
                    wmslpnSnItem.SN = DtSn.Rows[j]["serialno"].ToString();
                    //将两个产品加入List
                    list.Add(wmslpnSnItem);
                }
                //生成批次（装配、包装的箱号批次，系统车间的装箱批次）

                wmslpnItem.LPN = DtLot.Rows[i]["lotno"].ToString(); //大箱批次号
                wmslpnItem.LPNSize = DtLot.Rows[i]["qty"].ToString(); //大箱容量
                wmslpnItem.setNo = "";  //系统车间生成的才填
                wmslpnItem.remark = DtLot.Rows[i]["frerange"].ToString();
                wmslpnItem.module = DtLot.Rows[i]["machinetype"].ToString();
                wmslpnItem.packingName = ""; //包装规格
                wmslpnItem.SNItems = list;


                wmsLPNItemList.Add(wmslpnItem);

            }
            //生成头文件，包含工单信息
            WMSLPNHeader wmslpnHeader = new WMSLPNHeader();
            wmslpnHeader.action = "CREATE";   // "CREATE"/"DELETE"
            wmslpnHeader.LPNQty = DtLot.Rows.Count.ToString(); //批次数量，有几个大箱就是几
            wmslpnHeader.organizationCode = orgid;//组织代号
            if (dept != "SYS")
            {
                wmslpnHeader.sourceCode = "OTHERS";// "OTHERS"/"SYSTEM"     系统的装箱清单用SYSTEM，其他车间用OTHERS
                wmslpnHeader.moNumber = "";
            }
            else
            {
                wmslpnHeader.sourceCode = "SYSTEM";
                wmslpnHeader.moNumber = "";
            }
            wmslpnHeader.LPNItems = wmsLPNItemList;
            wmslpnHeader.workOrder = DtLot.Rows[0]["workno"].ToString();//工单号

            WMSLPNInputParamData wmslpnInputParamData = new WMSLPNInputParamData();
            wmslpnInputParamData.header = wmslpnHeader;

            WMSLPNRequestData wmsLPNRequestData = new WMSLPNRequestData();
            wmsLPNRequestData.wsinterface = wmslpnInputParamData;

            string batchNumn = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string method = "HYT_LPN_WS";
            string requestData = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(wmsLPNRequestData)));
            string responseData = string.Empty;
            bool result = WMSInterface.WMSInterfaceUtils.CallErpInterface(batchNumn, method, requestData, ref responseData);
            string s = "0";
            if (result)
                s = "1";
            else
                s = "0";
            string sqlupdate = " update lotnoinfo set IfSucess='" + s + "' where lotno='" + DtLot.Rows[0]["lotno"].ToString() + "' and org_id='" + orgid + "'";
            sqlupdate += " insert into WMSInterfaceLog(org_id, workno,lotno, batchNumn, method, requestData, result,opertype, uploaddate)";
            sqlupdate += "values('" + orgid + "','" + DtLot.Rows[0]["workno"].ToString() + "','" + DtLot.Rows[0]["lotno"].ToString() + "','" + batchNumn + "','" + method + "','" + responseData + "','" + s + "','" + wmslpnHeader.action + "',getdate())";
            DbAccess.ExecuteSql(sqlupdate);
        }
    }
}
