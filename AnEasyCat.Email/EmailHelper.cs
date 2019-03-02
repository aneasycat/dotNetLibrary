using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;

namespace AnEasyCat.Email
{
    public class EmailHelper
    {
        private string Sender = "";
        private SmtpClient smtp;
        public string Error = "";
        public EmailHelper(Models.Config config)
        {
            Sender = config.Sender;
            smtp = new SmtpClient(config.Host, config.Port)
            {
                Credentials = new NetworkCredential(config.Sender, config.Password)
            };
        }
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
        public bool Send(Models.Body body)
        {
            body.Sender = Sender;
            try
            {
                smtp.Send(body.MailMessage);
                return true;
            }
            catch(SmtpException ex)
            {
                Error = ex.Message.ToString();
                return false;
            }
        }
        public bool Send(MailMessage message)
        {
            try
            {
                smtp.Send(message);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public string Send(params Models.Body[] bodys)
        {
            string reStr = "";
            foreach(var body in bodys)
            {
                reStr += Send(body);
            }
            return reStr;
        }
        /// <summary>
        /// 异步发送邮件
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
        public bool SendAsync(Models.Body body)
        {
            body.Sender = Sender;
            try
            {
                smtp.SendAsync(body.MailMessage,null);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public string SendAsync(params Models.Body[] bodys)
        {
            string reStr = "";
            foreach (var body in bodys)
            {
                reStr += SendAsync(body);
            }
            return reStr;
        }
    }
}
