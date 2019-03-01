using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace AnEasyCat.dotNetLibrary
{
    public class EmailHelper
    {
        SmtpClient client;
        /// <summary>
        /// 初始化-默认端口25
        /// </summary>
        /// <param name="host">发件服务器地址</param>
        public EmailHelper(string host)
        {
            client = new SmtpClient(host, 25);
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="host">发件服务器地址</param>
        /// <param name="port">发件服务器端口</param>
        public EmailHelper(string host, int port)
        {
            client = new SmtpClient(host, port);
        }
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="Sender">发件人</param>
        /// <param name="SenderPassword">发件人邮箱密码</param>
        /// <param name="Receive">收件人</param>
        /// <param name="Title">标题</param>
        /// <param name="Body">内容</param>
        /// <returns>是否成功发送</returns>
        public bool Send(string Sender, string SenderPassword, string Receive, string Title, string Body)
        {
            try
            {
                MailMessage email = new MailMessage(Sender, Receive, Title, Body)
                {
                    IsBodyHtml = true
                };
                client.Credentials = new System.Net.NetworkCredential(Sender, SenderPassword);
                client.Send(email);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
