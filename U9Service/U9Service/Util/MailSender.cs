using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Web;

namespace U9Service.Util
{
    public class MailSender
    {
        /// <summary>
        /// 用于 SMTP 事务的主机的名称或 IP 地址
        /// </summary>
        public string Host { get; set; }
        /// <summary>
        /// 端口号
        /// </summary>
        public int Port { get; set; }
        /// <summary>
        /// 邮箱账号
        /// </summary>
        public string Email { get; set; }
        /// <summary>
        /// 邮箱密码
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// 发件人
        /// </summary>
        public MailAddress From { get; set; }
        /// <summary>
        /// 收件人集合
        /// </summary>
        public ArrayList To        {get;set;        }
        /// <summary>
        /// 抄送人集合
        /// </summary>
        public ArrayList CC { get; set; }
        /// <summary>
        /// 密送人集合
        /// </summary>
        public ArrayList Bcc { get; set; }
        /// <summary>
        /// 邮件标题
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// 邮件正文
        /// </summary>
        public string Body { get; set; }
        /// <summary>
        /// 邮件正文是否是Html格式，通过设置True/False来改变邮件正文格式，默认为True
        /// </summary>
        public string IsBodyHtml { get; set; }
        /// <summary>
        /// 标题编码格式
        /// </summary>
        public Encoding SubjectEncoding { get; set; }
        /// <summary>
        /// 邮件正文编码格式
        /// </summary>
        public Encoding BodyEncoding { get; set; }

        /// <summary>
        /// 发送邮件
        /// </summary>
        public void SendMail()
        {
            MailMessage mail = GetMail();
            SmtpClient client = new SmtpClient(this.Host, this.Port);//TODO:配置在web.config中
            //SmtpClient client = new SmtpClient("192.168.1.202");
            //sc.Credentials = new NetworkCredential("QuanZhenXiong6009000135", "qq123123");
            client.Credentials = new NetworkCredential(this.Email, this.Password);
            try
            {
                client.Send(mail);
            }
            catch (Exception ex)
            {
                //TODO:记录异常原因            
            }
        }
        /// <summary>
        /// 设置邮件信息
        /// </summary>
        /// <returns></returns>
        private MailMessage GetMail()
        {
            MailMessage mail = new MailMessage();
            if (!string.IsNullOrEmpty(this.From.Address))
            {
                mail.From = this.From;

            }
            else
            {
                throw new Exception("收件人不能为空！");
            }
            foreach (string item in this.To)
            {
                mail.To.Add(item);
            }
            foreach (string item in this.CC)
            {
                mail.CC.Add(item);
            }
            foreach (string item in this.Bcc)
            {
                mail.Bcc.Add(item);
            }
            if (string.IsNullOrEmpty(this.IsBodyHtml) || this.IsBodyHtml == "true")
            {
                mail.IsBodyHtml = true;
            }
            else if (this.IsBodyHtml == "false")
            {
                mail.IsBodyHtml = false;
            }
            else
            {
                throw new Exception("正文格式IsBodyHtml设置错误！");
            }
            mail.Subject = this.Subject;
            mail.SubjectEncoding = this.SubjectEncoding;
            mail.Body = this.Body;
            mail.BodyEncoding = this.BodyEncoding;
            return mail;
        }
    }
}