using System;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;

namespace _6_de_Mayo
{
    internal class Correo
    {
        private string smtpServer = "smtp.office365.com";  // Servidor SMTP para Office365
        private string fromEmail = "111473@alumnouninter.mx";     // Cambiar por tu correo de Office365
        private string fromPassword = "Gaming3AdunThoridas@";  // Cambiar por tu contraseña de Office365
        private string toEmail = "ecorrales@uninter.edu.mx";       // Correo del destinatario fijo

        public Correo() { } // Constructor sin parámetros

        public void EnviarCorreo(string asunto, string cuerpo, string rutaArchivo = null)
        {
            MailMessage mail = new MailMessage();
            SmtpClient smtpClient = new SmtpClient(smtpServer);

            try
            {
                mail.From = new MailAddress(fromEmail);
                mail.To.Add(toEmail);
                mail.Subject = asunto;
                mail.Body = cuerpo;
                mail.IsBodyHtml = true;

                // Adjuntar el archivo si se proporciona una ruta
                if (!string.IsNullOrEmpty(rutaArchivo) && File.Exists(rutaArchivo))
                {
                    mail.Attachments.Add(new Attachment(rutaArchivo, MediaTypeNames.Application.Octet));
                }

                smtpClient.Port = 587; // Puerto SMTP para Office365
                smtpClient.Credentials = new NetworkCredential(fromEmail, fromPassword);
                smtpClient.EnableSsl = true; // Habilitar SSL/TLS

                smtpClient.Send(mail);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al enviar el correo: " + ex.Message);
                EnviarCorreoError("Error al enviar correo", "Ocurrió un error al intentar enviar un correo: " + ex.Message);
            }
            finally
            {
                if (mail != null) mail.Dispose();
                if (smtpClient != null) smtpClient.Dispose();
            }
        }

        public void EnviarCorreoError(string asunto, string cuerpo)
        {
            MailMessage mail = new MailMessage();
            SmtpClient smtpClient = new SmtpClient(smtpServer);

            try
            {
                mail.From = new MailAddress(fromEmail);
                mail.To.Add(fromEmail); // Enviar el error al mismo remitente
                mail.Subject = asunto;
                mail.Body = cuerpo;
                mail.IsBodyHtml = true;

                smtpClient.Port = 587; // Puerto SMTP para Office365
                smtpClient.Credentials = new NetworkCredential(fromEmail, fromPassword);
                smtpClient.EnableSsl = true;

                smtpClient.Send(mail);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al enviar el correo de error: " + ex.Message);
            }
            finally
            {
                if (mail != null) mail.Dispose();
                if (smtpClient != null) smtpClient.Dispose();
            }
        }
    }
}