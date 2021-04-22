using Maticsoft.DBUtility;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using System.Web;
using System.Web.Services;
using U9Service.Util;

namespace U9Service
{
    /// <summary>
    /// U9Service 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    [System.Web.Script.Services.ScriptService]
    public class U9Service : System.Web.Services.WebService
    {
        [WebMethod]                
        public void MailTest()
        {
            try
            {
                DataSet ds = new DataSet();
                ds = DbHelperSQL.Query("exec sp_Auctus_MailToSupplier N''");
                if (ds.Tables[1].Rows.Count == 0)
                {
                    return;
                }
                TheadPar par = new TheadPar();
                par.ds = ds;
                par.UserName = "SSAdmin";
                Thread th = new Thread(new ParameterizedThreadStart(MailToSupplier));
                th.Start(par);
            }
            catch (Exception ex)
            {
                return;
            }
            return;
        }
        /// <summary>
        /// 邮件推送供应商欠料信息        
        /// </summary>
        /// <returns></returns>
        [WebMethod]
        public string MailToSupplier(string supplierID, string userName)
        {
            string result = "false";
            try
            {
                DataSet ds = new DataSet();
                ds = DbHelperSQL.Query("exec sp_Auctus_MailToSupplier N'" + supplierID + "'");
                if (ds.Tables[1].Rows.Count == 0)
                {
                    return "false";
                }
                TheadPar par = new TheadPar();
                par.ds = ds;
                par.UserName = userName;
                Thread th = new Thread(new ParameterizedThreadStart(MailToSupplier));
                th.Start(par);
                result = "true";
            }
            catch (Exception ex)
            {
                result = "false";
            }
            return result;
        }
        /// <summary>
        /// 8周欠料结果邮件发送功能
        /// </summary>
        /// <param name="obj"></param>
        public void MailToSupplier(object obj)
        {
            TheadPar par = (TheadPar)obj;
            string userName = par.UserName;
            //邮箱信息设置
            MailSender mailSender = new MailSender();
            mailSender.Email = "sys_sup@auctus.cn";
            mailSender.Password = "Qwelsy@1234";
            mailSender.Host = "192.168.1.1";
            mailSender.Port = 25;
            mailSender.IsBodyHtml = "true";
            mailSender.From = new MailAddress("sys_sup@auctus.cn", "深圳力同芯科技发展有限公司");
            mailSender.To = new ArrayList();
            mailSender.CC = new ArrayList();
            mailSender.Bcc = new ArrayList();
            //数据源
            DataSet ds = par.ds;
            //Table0为无邮箱供应商
            if (ds.Tables[0].Rows.Count > 0)
            {
                GetNoneEmaiContent(ds.Tables[0], ref mailSender, userName);
            }
            //Table1为有邮箱供应商
            if (ds.Tables[1].Rows.Count > 0)
            {
                GetEmaiContent(ds.Tables[1], ref mailSender, userName);
            }
        }
        /// <summary>
        /// 无邮箱供应商邮件内容
        /// </summary>
        /// <param name="dt">无邮箱信息供应商集合</param>
        /// <param name="liSql"></param>
        /// <param name="mailSender">邮件发送对象</param>
        private void GetNoneEmaiContent(DataTable dt, ref MailSender mailSender, string userName)
        {
            List<string> liSql = new List<string>();//邮件发送日志sql
            //邮件样式
            string style = "<style>table,table tr th, table tr td { border:2px solid #cecece; } table {text-align: center; border-collapse: collapse; padding:2px;}</style>";
            string noneEmailSup = "";//邮件内容
            noneEmailSup += style + "<H2>Dear All,</H2><H2></br>&nbsp;&nbsp;以下是无默认邮箱供应商，请相关人员知悉。谢谢！</H2>";
            noneEmailSup += "<table><tr bgcolor='#cae7fc'><td>供应商名称</td><td>供应商执行采购</td></tr>";
            mailSender.Subject = "无邮箱供应商列表";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                noneEmailSup += "<tr><td>";
                noneEmailSup += dt.Rows[i]["Supplier"].ToString();
                noneEmailSup += "</td>";
                noneEmailSup += "<td>";
                noneEmailSup += dt.Rows[i]["Purchaser"].ToString();
                noneEmailSup += "</td></tr>";
            }
            noneEmailSup += "</table>";
            mailSender.Body = noneEmailSup;
            mailSender.To=new ArrayList(dt.Rows[0]["Email"].ToString().TrimEnd(';').Split(';'));
            string sql = "";
            string toContent = ArrayListToStr(mailSender.To);
            string ccContent = ArrayListToStr(mailSender.CC);
            try
            {
                mailSender.SendMail();
                sql = "insert into Auctus_MailLog values('" + userName + "','" + mailSender.Subject + "','" + mailSender.Body.Replace("'", "''") + "','" + mailSender.From + "','" + toContent + "','" + ccContent + "',GetDate(),1,'')";
            }
            catch (Exception ex)
            {
                sql = "insert into Auctus_MailLog values('" + userName + "','" + mailSender.Subject + "','" + mailSender.Body.Replace("'", "''") + "','" + mailSender.From + "','" + toContent + "','" + ccContent + "',GetDate(),1,'" + ex.Message.ToString().Replace("'", "''") + "')";
            }
            liSql.Add(sql);
            DbHelperSQL.ExecuteSqlTran(liSql);//保存邮件发送信息
        }
        /// <summary>
        /// 发送欠料信息邮件到供应商
        /// </summary>
        /// <param name="dt">邮件信息集合</param>
        /// <param name="liSql">邮件发送日志sql</param>
        /// <param name="mailSender">邮件发送对象</param>
        private void GetEmaiContent(DataTable dt, ref MailSender mailSender, string userName)
        {
            List<string> liSql = new List<string>();//邮件发送日志sql
            string strBody = "";
            string style = "<style>table,table tr th, table tr td { border:2px solid #cecece; } table {text-align: center; border-collapse: collapse; padding:2px;}</style>";
            int index = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["OrderNo"]) + 1;
            //邮件发送信息设置              
            mailSender.Subject = "供应商需求交货计划";
            //保存邮件发送记录                
            for (int i = 1; i < index; i++)
            {
                string emailContent = "";
                DataRow[] dr = dt.Select("OrderNo=" + i.ToString());
                if (dr.Length > 0)
                {
                    strBody = "<H2>" + dr[0]["Supplier"].ToString() + "：</H2><H2></br>&nbsp;&nbsp;如下未来8周需求计划，供生产备货安排!</H2><h2>&nbsp;&nbsp;请务必达成交期，如有问题请及时反馈，谢谢配合与支持！！ </h2>";
                    strBody += "<span style='font-weight:bold;'>备注：</span></br><span>电子物料要求交货数量：(紧急欠料+第一周+第二周+第三周)-已交在检</span></br><span>结构物料要求交货数量：(紧急欠料+第一周)-已交在检</span></br><span>包材/配件物料要求交货数量：(紧急欠料+第一周)-已交在检</span></br><span>**预测数量仅供参考提前做生产安排，实际按计划数量如期交货**</span>";
                    //strBody = "<h2>&nbsp;&nbsp;请验证数据是否有问题，谢谢！</h2>";
                    emailContent += style + strBody;
                    emailContent += "<table>";
                    emailContent += @"<tr bgcolor='#cae7fc'>  <td nowrap = 'nowrap'>供应商</td><td colspan='22'>" + dr[0]["Supplier"].ToString() + "</td></tr>";
                    emailContent += "<tr bgcolor = '#cae7fc' ><td nowrap = 'nowrap'> 采购员 </td ><td colspan = '22' >" + dr[0]["Purchaser"].ToString() + "</td ></tr > ";
                    emailContent += @"<tr bgcolor = '#cae7fc' > 
             <td nowrap = 'nowrap' rowspan = '3' > 分类 </td >    
             <td nowrap = 'nowrap' rowspan = '3' > 料号 </td >    
                <td nowrap = 'nowrap'style = 'min-width:160px;' rowspan = '3' > 品名 </td >         
                     <td nowrap = 'nowrap' style = 'min-width:300px;' rowspan = '3' > 规格 </td >              
                          <td nowrap = 'nowrap' rowspan = '3' > 已交在检 </td >                 
                          <td nowrap = 'nowrap' rowspan = '3' > 紧急欠料 </td >                 
                             <td colspan='2'  nowrap = 'nowrap' > 第一周 </td >         <td colspan='2'  nowrap = 'nowrap' > 第二周 </td >            <td colspan='2'  nowrap = 'nowrap' > 第三周 </td >                                  
                                              <td colspan='2' nowrap = 'nowrap' > 第四周 </td >            <td colspan='2' nowrap = 'nowrap' > 第五周 </td >            <td colspan='2' nowrap = 'nowrap' > 第六周 </td >                                                  
                                                              <td colspan='2' nowrap = 'nowrap' > 第七周 </td >            <td colspan='2' nowrap = 'nowrap' > 第八周 </td >         <td nowrap = 'nowrap' rowspan = '3' > 8周欠料数量 </td >   </tr >";
                }
                emailContent += "<tr bgcolor='#cae7fc'>";
                DateTime monday = GetMonday(DateTime.Now);
                for (int w = 0; w < 8; w++)
                {
                    //emailContent += "<td colspan='2' nowrap='nowrap'>" + monday.AddDays(w * 7).ToString("MM-dd") + "~" + monday.AddDays((w + 1) * 7 - 1).ToString("MM-dd")+"</td>";
                    emailContent += "<td colspan='2' nowrap='nowrap'>" + monday.AddDays(w * 7).ToString("MM-dd") + "</td>";
                }
                emailContent += "</tr>";

                emailContent += @"<tr bgcolor='#cae7fc'>
<td nowrap='nowrap' colspan='2'>计划交货数量</td>
<td nowrap='nowrap' colspan='2' >计划交货数量</td>
<td nowrap='nowrap' colspan='2'>计划交货数量</td>
<td nowrap='nowrap'>计划交货数量</td>
<td nowrap='nowrap'>预测交货数量</td>
<td nowrap='nowrap'>计划交货数量</td>
<td nowrap='nowrap'>预测交货数量</td>
<td nowrap='nowrap'>计划交货数量</td>
<td nowrap='nowrap'>预测交货数量</td>
<td nowrap='nowrap'>计划交货数量</td>
<td nowrap='nowrap'>预测交货数量</td>
<td nowrap='nowrap'>计划交货数量</td>
<td nowrap='nowrap'>预测交货数量</td>
<tr>";
                ArrayList to = new ArrayList();
                ArrayList cc = new ArrayList();
                for (int j = 0; j < dr.Length; j++)
                {
                    emailContent += "<tr>";
                    //拼接邮件内容
                    emailContent += "<td>" + dr[j]["MRPCategory"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["Code"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["Name"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["SPECS"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["RcvQty"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w0"].ToString() + "</td>";
                    emailContent += "<td colspan='2'>" + dr[j]["w1"].ToString() + "</td>";
                    emailContent += "<td colspan='2'>" + dr[j]["w2"].ToString() + "</td>";
                    emailContent += "<td colspan='2'>" + dr[j]["w3"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w4"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w42"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w5"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w52"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w6"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w62"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w7"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w72"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w8"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["w82"].ToString() + "</td>";
                    emailContent += "<td>" + dr[j]["Total"].ToString() + "</td>";
                    emailContent += "</tr>";
                    //拼接收件人
                    if (dr[j]["Email"] != null && dr[j]["Email"].ToString() != "")
                    {
                        string[] arrTo = dr[j]["Email"].ToString().TrimEnd(';').Split(';');
                        for (int m = 0; m < arrTo.Length; m++)
                        {
                            if (!to.Contains(arrTo[m]))
                            {
                                to.Add(arrTo[m]);
                            }
                        }
                    }
                    //拼接抄送人
                    if (dr[j]["CC"] != null && dr[j]["CC"].ToString() != "")
                    {
                        string[] arrCC = dr[j]["cc"].ToString().TrimEnd(';').Split(';');
                        for (int m = 0; m < arrCC.Length; m++)
                        {
                            if (!cc.Contains(arrCC[m]))
                            {
                                cc.Add(arrCC[m]);
                            }
                        }
                    }
                }
                emailContent += "</table>";
                mailSender.To = to;
                mailSender.CC = cc;
                mailSender.Body = emailContent;
                string toContent = ArrayListToStr(mailSender.To);//拼接收件人字符串
                string ccContent = ArrayListToStr(mailSender.CC);//拼接抄送人字符串
                string sql = "";
                try
                {
                    mailSender.SendMail();
                    sql = "insert into Auctus_MailLog values('" + userName + "','" + mailSender.Subject + "','" + mailSender.Body.Replace("'", "''") + "','" + mailSender.From + "','" + toContent + "','" + ccContent + "',GetDate(),1,'')";
                }
                catch (Exception ex)
                {
                    sql = "insert into Auctus_MailLog values('" + userName + "','" + mailSender.Subject + "','" + mailSender.Body.Replace("'", "''") + "','" + mailSender.From + "','" + toContent + "','" + ccContent + "',GetDate(),1,'" + ex.Message.ToString().Replace("'", "''") + "')";
                }
                liSql.Add(sql);
                //邮件服务器扛不住，每20s发6封邮件
                if (i % 6 == 0)
                {
                    Thread.Sleep(20000);
                }
            }
            DbHelperSQL.ExecuteSqlTran(liSql);//保存邮件发送信息
        }

        /// <summary>
        /// ArrayList集合拼接字符串，以分号";"连接
        /// </summary>
        /// <param name="CC"></param>
        /// <returns></returns>
        private string ArrayListToStr(ArrayList arrList)
        {
            string str = "";
            for (int t = 0; t < arrList.Count; t++)
            {
                str += arrList[t].ToString() + ";";
            }
            return str;
        }
        /// <summary>
        /// 获取当前周的星期一
        /// </summary>
        /// <param name="Date"></param>
        /// <returns></returns>
        public DateTime GetMonday(DateTime Date)
        {

            if (Date.DayOfWeek == 0)
            {
                return Date.AddDays(-6);
            }
            else
            {
                return Date.AddDays((Convert.ToInt32(Date.DayOfWeek) - 1) * (-1));
            }
        }

    }
}

