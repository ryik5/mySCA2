using System;
using System.Linq;
using System.Threading.Tasks;
using MimeKit;

namespace ASTA.Classes
{
   sealed class MailSender
    {
        MailServer _mailServer;
        MailUser _from;
        MailUser _to;
        public string Status { get; private set; }

        public MailSender() { }

        public MailSender(MailServer mailServer)
        {
            _mailServer = mailServer;
        }

        public void SetFrom(MailUser from)
        {
            if (string.IsNullOrWhiteSpace(from.GetEmail()) || !from.GetEmail().Contains('@'))
            {
                throw new ArgumentException("no 'From' address provided!");
            }
            _from = from;
        }

        public void SetTo(MailUser to)
        {
            if (string.IsNullOrWhiteSpace(to.GetEmail()) || !to.GetEmail().Contains('@'))
            {
                throw new ArgumentException("no 'To' address provided!");
            }
            Status = string.Empty;
            _to = to;
        }

        public void SetServer(MailServer mailServer)
        {
            if (string.IsNullOrWhiteSpace(mailServer.GetName()))
            {
                throw new ArgumentException("no Server's name provided!");
            }

            if (mailServer.GetPort() < 0)
            {
                throw new ArgumentException("no Server's port provided!");
            }

            _mailServer = mailServer;
        }

        public async Task SendEmailAsync(string subject, BodyBuilder bodyOfMessage)
        {
            if ((string.IsNullOrEmpty(_mailServer.ToString()) ||
                string.IsNullOrEmpty(_from.ToString()) ||
                string.IsNullOrEmpty(_to.ToString())) && !_from.GetEmail().Contains('@') && !_to.GetEmail().Contains('@'))
            {
                throw new ArgumentException("no needed data provided!");
            }

            MailKit.Net.Smtp.SmtpClient client;
            MimeMessage emailMessage = new MimeMessage();
            emailMessage.From.Add(new MailboxAddress(_from.GetName(), _from.GetEmail()));
            emailMessage.To.Add(new MailboxAddress(_to.GetEmail()));
            emailMessage.Subject = subject;
            emailMessage.Body = bodyOfMessage.ToMessageBody();
            // emailMessage.Headers[HeaderId.DispositionNotificationTo] = new MailboxAddress(_from.GetName(), _from.GetEmail()).ToString(true); //request a notification when the message is read by the user

            using (client = new MailKit.Net.Smtp.SmtpClient(new MailKit.ProtocolLogger("smtp.log", false))) //
            {
                client.Timeout = (int)TimeSpan.FromSeconds(10).TotalMilliseconds;
                try
                {
                    // For demo-purposes, accept all SSL certificates (in case the server supports STARTTLS)
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    await client.ConnectAsync(_mailServer.GetName(), _mailServer.GetPort(), false).ConfigureAwait(false);
                    client.AuthenticationMechanisms.Remove("XOAUTH2");
                    // Note: only needed if the SMTP server requires authentication
                    // client.AuthenticateAsync(_from.GetEmail(),"password").ConfigureAwait(false);
                    client.MessageSent += OnMessageSent;
                    await client.SendAsync(emailMessage).ConfigureAwait(false);  //client.SendAsync(emailMessage).ConfigureAwait(false)
                }
                catch (Exception e)
                {
                    Status = e.Message;
                    //Status = _to.GetEmail() + " - " + e.Message;
                }

                await client.DisconnectAsync(true).ConfigureAwait(false);
            }
            client.Dispose();
        }

        void OnMessageSent(object sender, MailKit.MessageSentEventArgs e)
        {
            Status = " успешно";
            // Status = e.Message.To + " - OK";
        }

        /*
        System.Net.Mail.LinkedResource mailLogo;
        private void SetMail()
        {
            //e-mail logo
            //convert embedded resources into memory stream to attach at an email
            Bitmap b = new Bitmap(Properties.Resources.LogoRYIK, new Size(50, 50));
            ImageConverter ic = new ImageConverter();
            Byte[] ba = (Byte[])ic.ConvertTo(b, typeof(Byte[]));
            System.IO.MemoryStream logo = new System.IO.MemoryStream(ba);
            mailLogo = new System.Net.Mail.LinkedResource(logo, "image/jpeg");
            mailLogo.ContentId = Guid.NewGuid().ToString(); //myAppLogo for email's reports
        }

        //Compose and send e-mail
        private static void SendEmail(string to, string period, string department, string pathToFile, string messageAfterPicture)
        {
            StartStopTimer timer1sec = new StartStopTimer(1);
            StartStopTimer timer10sec = new StartStopTimer(10);
            //    mailStopSent = false;
            // string startupPath = AppDomain.CurrentDomain.RelativeSearchPath;
            // string path = System.IO.Path.Combine(startupPath, "HtmlTemplates", "NotifyTemplate.html");
            // string body = System.IO.File.ReadAllText(path);
            // адрес smtp-сервера и порт, с которого будем отправлять письмо
            using (System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(mailServer, mailServerSMTPPort))
            {
                smtpClient.EnableSsl = false; // I got error with "true"
                smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Timeout = 50000;

                {
                     logger.Trace("SendEmail: Try to send From:" + mailsOfSenderOfName + "| To:" + to + "| " + period + "| " + department + "|" + pathToFile + "|" + messageAfterPicture);
                     logger.Trace("SendEmail: file=" + (pathToFile.Length > 0) + ", admin=" + to.Equals(mailJobReportsOfNameOfReceiver));

                    // создаем объект сообщения
                    using (System.Net.Mail.MailMessage newMail = new System.Net.Mail.MailMessage())
                    {
                        // письмо представляет код html
                        newMail.IsBodyHtml = true;
                        newMail.BodyEncoding = Encoding.UTF8;

                        if (pathToFile.Length > 0)
                        {
                            newMail.AlternateViews.Add(GetStandartReportMessageOfBody(period, department, messageAfterPicture));
                            // тема письма
                            newMail.Subject = "Отчет по посещаемости за период: " + period;
                            // отправитель - устанавливаем адрес и отображаемое в письме имя
                            newMail.From = new System.Net.Mail.MailAddress(mailsOfSenderOfName, NAME_OF_SENDER_REPORTS);
                            newMail.ReplyToList.Add(mailsOfSenderOfName);

                            // отправитель - устанавливаем адрес и отображаемое в письме имя
                            //   newMail.ReplyToList.Add(mailsOfSenderOfName);                  
                            // newMail.Bcc.Add("user@mail.com.ua");      // Скрытая копия

                            // кому отправляем
                            newMail.To.Add(new System.Net.Mail.MailAddress(to));
                            // Добавляем проверку успешности отправки
                            smtpClient.SendCompleted += new System.Net.Mail.SendCompletedEventHandler(SendCompletedCallback);
                            // добавляем вложение
                            newMail.Attachments.Add(new System.Net.Mail.Attachment(pathToFile));

                            // логин и пароль
                            smtpClient.Credentials = new System.Net.NetworkCredential(mailsOfSenderOfName.Split('@')[0], "");

                            // отправка письма
                            int attempts = 5;
                            bool sending = true;
                            while (sending)
                            {
                                try
                                {
                                    // async sending method sometimes has a problem with sending emails
                                    // a default MS Exch Server blocks the action of mass sending emails
                                    // smtpClient.SendAsync(newMail, userState);
                                    smtpClient.Send(newMail);

                                    resultOfSendingReports.Add(new Mailing
                                    {
                                        _recipient = to,
                                        _nameReport = pathToFile,
                                        _descriptionReport = department,
                                        _period = period,
                                        _status = "Ok"
                                    });
                                    logger.Info("SendEmail:  To:" + to + "|" + department + "|" + period + " |" + pathToFile + " - Ok");

                                    sending = false;
                                    timer1sec.WaitTime();
                                }
                                catch (Exception expt) // "Error sending the email"
                                {
                                    logger.Error("SendEmail, Error to send: " + to + " |by: " + mailServer + ":" + mailServerSMTPPort + " " + expt.Message);

                                    resultOfSendingReports.Add(new Mailing
                                    {
                                        _recipient = to,
                                        _nameReport = pathToFile,
                                        _descriptionReport = mailServer + ":" + mailServerSMTPPort,
                                        _period = period,
                                        _status = "Error: " + expt.Message
                                    });

                                    if (expt.Message.Contains("Unknown user"))
                                    {
                                        smtpClient.SendAsyncCancel(); //stop sending
                                        sending = false;
                                    }
                                    else
                                    {
                                        timer10sec.WaitTime();
                                    }
                                }
                                finally
                                {
                                    attempts--;
                                }

                                if (attempts < 0)
                                {
                                    smtpClient.SendAsyncCancel(); //stop sending
                                    sending = false;
                                    resultOfSendingReports.Add(new Mailing
                                    {
                                        _recipient = to,
                                        _nameReport = pathToFile,
                                        _descriptionReport = mailServer + ":" + mailServerSMTPPort,
                                        _period = period,
                                        _status = "Email has not been sent"
                                    });
                                }
                                if (mailStopSent == false)
                                {
                                    smtpClient.SendAsyncCancel();
                                }
                            }
                            timer10sec.WaitTime();
                        }
                    }
                }
            }
        }

        private static void SendCompletedCallback(object sender, System.ComponentModel.AsyncCompletedEventArgs e)     //for async sending
        {
            //Get the Original MailMessage object
            System.Net.Mail.MailMessage mail = (System.Net.Mail.MailMessage)e.UserState;
            //write out the subject
            string subject = mail.Subject.ToString();
            string recepient = mail.To.ToString();

            // Get the unique identifier for this asynchronous operation.
            // String token = (string)e.UserState;

            //  if (e.Cancelled)
            //   { logger.Warn("Send canceled " + " to "+ recepient+"|"+ subject); }
            if (e.Error != null)
            { logger.Warn("Send error to " + recepient + "|" + subject + "|" + e.Error.ToString()); }
            else
            {  logger.Trace("Message sent to " + recepient + "|" + subject); }
            mailStopSent = true;
        }


        private static System.Net.Mail.AlternateView GetStandartReportMessageOfBody(string period, string department, string messageAfterPicture)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(@"<style type = 'text/css'> A {text - decoration: none;}</ style >");
            sb.Append(@"<p><font size='3' color='black' face='Arial'>Здравствуйте,</p>Во вложении «Отчет по учету рабочего времени сотрудников».<p>");

            sb.Append(@"<b>Период: </b>");
            sb.Append(period);

            sb.Append(@"<br/><b>Подразделение: </b>'");
            sb.Append(department);
            sb.Append(@"'<br/><p>Уважаемые руководители,</p><p>согласно Приказу ГК АИС «О функционировании процессов кадрового делопроизводства»,<br/><br/>");
            sb.Append(@"<b>Внесите,</b> пожалуйста, <b>до конца текущего месяца</b> по сотрудникам подразделения ");
            sb.Append(@"информацию о командировках, больничных, отпусках, прогулах и т.п. <b>на сайт</b> www.ais .<br/><br/>");
            sb.Append(@"Руководители <b>подразделений</b> ЦОК, <b>не отображающихся на сайте,<br/>вышлите, </b>пожалуйста, <b>Табель</b> учета рабочего времени<br/>");
            sb.Append(@"<b>в отдел компенсаций и льгот до последнего рабочего дня месяца.</b><br/></p>");
            sb.Append(@"<font size='3' color='black' face='Arial'>С, Уважением,<br/>");
            sb.Append(NAME_OF_SENDER_REPORTS);
            sb.Append(@"</font><br/><br/><font size='2' color='black' face='Arial'><i>");
            sb.Append(@"Данное сообщение и отчет созданы автоматически<br/>программой по учету рабочего времени сотрудников.");
            sb.Append(@"</i></font><br/><font size='1' color='red' face='Arial'><br/>");
            sb.Append(DateTime.Now.ToYYYYMMDDHHMM());
            sb.Append(@"</font></p>");
            sb.Append(@"<hr><img alt='ASTA' src='cid:");
            sb.Append(mailLogo.ContentId);
            sb.Append(@"'/><br/><a href='mailto:ryik.yuri@gmail.com'>_</a>");

            System.Net.Mail.AlternateView alternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(sb.ToString(), null, System.Net.Mime.MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(mailLogo);
            return alternateView;
        }

        private static System.Net.Mail.AlternateView GetAdminReportMessageOfBody(string period, List<Mailing> reportOfResultSending)
        {
            int order = 0;
            StringBuilder sb = new StringBuilder();
            sb.Append(@"<style type = 'text/css'> A {text - decoration: none;}</ style >");
            sb.Append(@"<p><font size='4' color='black' face='Arial'>Здравствуйте,</p><p>Результаты отправки отчетов</p>");

            sb.Append(@"<b>Дата: </b>");
            sb.Append(period);
            sb.Append(@"<p><b>Результат отправки " + reportOfResultSending.Count + @" отчета/отчетов:</b></p><p>");

            foreach (var mailing in reportOfResultSending)
            {
                order += 1;
                if (
                    mailing._status.ToLower().Contains("error") || mailing._status.ToLower().Contains("not been sent")
                    )
                {
                    sb.Append(@"<font size='2' color='red' face='Arial'>");
                    sb.Append(
                        order + @". Получатель: " + mailing._recipient +
                        @", " + mailing._descriptionReport +
                        @", " + mailing._status + @"</font><br/>");
                }
                else
                {
                    sb.Append(@"<font size='2' color='black' face='Arial'>");
                    sb.Append(
                        order + @". Получатель: " + mailing._recipient +
                        @", отчет по группе " + mailing._descriptionReport +
                        @", доставлен: " + mailing._status + @" </font><br/>");
                }
            }

            sb.Append(@"</p><font size='2' color='black' face='Arial'>С, Уважением,<br/>");
            sb.Append(NAME_OF_SENDER_REPORTS);
            sb.Append(@"</font><br/><br/><font size='1' color='black' face='Arial'><i>");
            sb.Append(@"Данное сообщение и отчет созданы автоматически<br/>программой по учету рабочего времени сотрудников.");
            sb.Append(@"</i></font><br/><font size='1' color='red' face='Arial'><br/>");
            sb.Append(DateTime.Now.ToYYYYMMDDHHMM());
            sb.Append(@"</font></p>");
            sb.Append(@"<hr><img alt='ASTA' src='cid:");
            sb.Append(mailLogo.ContentId);
            sb.Append(@"'/><br/><a href='mailto:ryik.yuri@gmail.com'>_</a>");

            System.Net.Mail.AlternateView alternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(sb.ToString(), null, System.Net.Mime.MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(mailLogo);
            return alternateView;
        }

        */
    }
}
