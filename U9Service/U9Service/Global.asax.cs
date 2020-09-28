using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Security;
using System.Web.SessionState;

namespace U9Service
{
    public class Global : System.Web.HttpApplication
    {
        protected void Application_Start(object sender, EventArgs e)
        {
            Thread thTimer = new Thread(new ThreadStart(ExecuteService));
            thTimer.Start();
            System.Timers.Timer cronJob = new System.Timers.Timer(1000);
            cronJob.Elapsed += new System.Timers.ElapsedEventHandler(SendMailByHour);
            cronJob.Enabled = true;
            cronJob.AutoReset = true;
            cronJob.Start();
        }
        protected void Application_End(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(2000);
            // 发送请求打开程序。
            string url = "http://192.168.1.226:9090/U9Service/WebForm1.aspx";
            //string url = "http://localhost:11614/WebForm1.aspx";
            
            System.Net.HttpWebRequest myHttpWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            System.Net.HttpWebResponse myHttpWebResponse = (System.Net.HttpWebResponse)myHttpWebRequest.GetResponse();
            System.IO.Stream receiveStream = myHttpWebResponse.GetResponseStream();//得到回写的字节流
        }

        private void ExecuteService()
        {

        }
        private void SendMailByHour(object sender, System.Timers.ElapsedEventArgs e)
        {
            int mint = e.SignalTime.Minute;
            int second = e.SignalTime.Second;
            int hour = e.SignalTime.Hour;
            // 星期一9点
            // string s = e.SignalTime.DayOfWeek.ToString();

            // 每天9点发送测试邮件
            // e.SignalTime.DayOfWeek == DayOfWeek.Monday&&
            if (e.SignalTime.DayOfWeek == DayOfWeek.Monday && hour == 9 && mint== 0 && second == 0)
            {
                string url = "http://192.168.1.226:9090/U9Service/U9Service.asmx/MailTest";
                //string url = "http://localhost:11614/U9Service.asmx/MailTest";                
                System.Net.HttpWebRequest myHttpWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);                
                System.Net.HttpWebResponse myHttpWebResponse = (System.Net.HttpWebResponse)myHttpWebRequest.GetResponse();
                System.IO.Stream receiveStream = myHttpWebResponse.GetResponseStream();//得到回写的字节流
            }
        }
    }
}