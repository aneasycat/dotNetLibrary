using System.Collections.Generic;
using System.Text;
using System.Net.Mail;

namespace AnEasyCat.Email
{
    public class Models
    {
        public class Config
        {
            /// <summary>
            /// smtp服务器地址
            /// </summary>
            public string Host { get; set; }
            /// <summary>
            /// smtp服务器端口
            /// </summary>
            public int Port { get; set; } = 25;
            /// <summary>
            /// 发件人邮箱账号
            /// </summary>
            public string Sender { get; set; }
            /// <summary>
            /// 发件人邮箱密码
            /// </summary>
            public string Password { get; set; }
        }
        public class Body
        {
            /// <summary>
            /// 发件人显示名
            /// </summary>
            public string SenderName { get; set; }
            /// <summary>
            /// 收件人
            /// </summary>
            public string[] Recipients { get; set; }
            /// <summary>
            /// 邮件标题
            /// </summary>
            public string Title { get; set; }
            /// <summary>
            /// 邮件优先级
            /// </summary>
            public MailPriority Priority { get; set; } = MailPriority.Normal;
            /// <summary>
            /// 邮件内容
            /// </summary>
            public string Content { get; set; }
            /// <summary>
            /// 邮件内容是否是html格式
            /// </summary>
            public bool IsHtml { get; set; } = false;
            /// <summary>
            /// 邮件附件
            /// </summary>
            public IList<Attachment> Attachments { get; set; }
            /// <summary>
            /// 邮件编码
            /// </summary>
            public Encoding Encoding { get; set; } = Encoding.UTF8;
            private MailMessage message = new MailMessage();
            public MailMessage MailMessage
            {
                get
                {
                    message.Subject = Title;
                    message.Priority= Priority;
                    message.Body = Content;
                    message.IsBodyHtml = IsHtml;
                    message.From = new MailAddress(Sender, SenderName, Encoding);
                    message.BodyEncoding = message.SubjectEncoding = Encoding;

                    if (Recipients != null && Recipients.Length > 0)
                        foreach (var recipient in Recipients)
                        {
                            message.To.Add(recipient);
                        }
                    if (Attachments != null && Attachments.Count > 0)
                        foreach (var attachment in Attachments)
                        {
                            message.Attachments.Add(attachment);
                        }
                    return message;
                }
            }
            /// <summary>
            /// 发件人地址-不赋值
            /// </summary>
            public string Sender { get; set; }
        }
    }
}
