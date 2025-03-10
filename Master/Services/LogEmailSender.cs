using System;
using System.Net.Mail;
using System.Threading.Tasks;

namespace CratePilotSystemWS.Services.Email
{
    public class EmailServices
    {
        public static async Task<(bool success, string msg)> SendTestEmail()
        {
            string senderEmail = "PEN_EST1C@jabil.com"; 
            string receiverEmail = "SyedFaizAlhady_SyedAhmadAlhady@jabil.com"; 
            string subject = "Test Email Connection";
            
            // ✅ Updated logo URL
            string logoUrl = "http://mypenm0iesvr02:5000/images/logo.png";

            // 📩 Email Body (HTML)
            string body = $@"
                <div style='font-family: Arial, sans-serif; text-align: center; padding: 20px;'>
                    <img src='{logoUrl}' alt='Jabil Logo' style='width: 150px; margin-bottom: 20px;'>
                    <h2 style='color: #333;'>Email Test Successful</h2>
                    <p>This is a test email from the <strong>EST1C System</strong>.</p>
                    <p><strong>Sent at:</strong> {DateTime.Now:dddd, dd MMMM yyyy HH:mm:ss}</p>
                    <hr style='border: 0; height: 1px; background: #ccc; margin: 20px 0;'>
                    <p style='font-size: 12px; color: #777;'>
                        <em>This is an auto-generated message. Please do not reply to this email.</em>
                    </p>
                </div>";

            int retryCount = 0;
            const int maxRetries = 2;
            string lastErrorMessage = string.Empty;
            string[] smtpHosts = { "MYPENM0IESVR02.corp.JABIL.ORG", "CORIMC04.corp.jabil.org" };

            do
            {
                try
                {
                    using MailMessage message = new();
                    using SmtpClient smtp = new();

                    message.From = new MailAddress(senderEmail, "PEN EST1C System");
                    message.To.Add(new MailAddress(receiverEmail));

                    message.Subject = subject;
                    message.IsBodyHtml = true; // ✅ Enable HTML format
                    message.Body = body;

                    smtp.Port = 25;
                    smtp.Host = smtpHosts[retryCount];
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

                    await smtp.SendMailAsync(message);
                    return (true, $"Email successfully sent to {receiverEmail}");
                }
                catch (Exception ex)
                {
                    lastErrorMessage = ex.Message;
                    retryCount++;
                }
            } while (retryCount < maxRetries);

            return (false, $"Email failed after {maxRetries} retries. Last error: {lastErrorMessage}");
        }
    }
}
