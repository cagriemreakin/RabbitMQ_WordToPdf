using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;
using System;
using System.IO;
using System.Net.Mail;
using System.Text;

namespace WordToPdfConverter.Consumer
{
    class Program
    {
        static void Main(string[] args)
        {

            var factory = new ConnectionFactory();
            factory.Uri = new Uri("url");
            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare(exchange:"convertExchange",ExchangeType.Direct,durable:true,autoDelete:false,null);
                    channel.QueueBind("File", "convertExchange", routingKey:"WordToPdf",null);

                    channel.BasicQos(0,1,false);

                    var consumer = new EventingBasicConsumer(channel);

                    channel.BasicConsume("File",false,consumer);
                    bool result = false;
                    consumer.Received += (model,ea)=>{

                        try
                        {
                            Console.WriteLine("Kuyrukran mesaj alındı,işleniyor");

                            Document doc = new Document();
                            string deserializeString = Encoding.UTF8.GetString(ea.Body.ToArray());
                            MessagWordToPdf messagWordToPdf = JsonConvert.DeserializeObject<MessagWordToPdf>(deserializeString);

                            doc.LoadFromStream(new MemoryStream(messagWordToPdf.WordByte), FileFormat.Docx2013);

                            using (MemoryStream ms =  new MemoryStream())
                            {
                               doc.SaveToStream(ms,FileFormat.PDF);
                               result = EmailSend(messagWordToPdf.Email, ms, messagWordToPdf.FileName);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Hata meydana geldi" + ex.InnerException);
                        }

                        if (result)
                        {
                            Console.WriteLine("Kuyrukta mesaj başarıyla işlendi");
                            channel.BasicAck(ea.DeliveryTag,false);
                        }
                        else
                        {
                            Console.WriteLine("Hata meydana geldi");

                        }
                    };
                    Console.WriteLine("Çıkmak için tıklayınız");
                    Console.ReadLine();
                }
            }


        }



        public static bool EmailSend( string email, MemoryStream ms,string fileName)
        {
            try
            {
                ms.Position = 0;
                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);

                Attachment attachment = new Attachment(ms, ct);
                attachment.ContentDisposition.FileName = $"{fileName}.pdf";

                MailMessage mailMessage = new MailMessage
                {
                    From = new MailAddress("sender-email-adres")
                };

                mailMessage.To.Add(email);
                mailMessage.Subject = "Rabbit Mq Word To Pdf Conversion";
                mailMessage.Body = "Pdf dosyanız ektedir.";
                mailMessage.IsBodyHtml = true;
                mailMessage.Attachments.Add(attachment);

                SmtpClient smtpClient = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                };
                smtpClient.UseDefaultCredentials = false;
                smtpClient.EnableSsl = true;
                smtpClient.Credentials = new System.Net.NetworkCredential("sender-email-adress", "password");

                smtpClient.Send(mailMessage);
                Console.WriteLine($"Email adresine {email} gönderilmiştir.");
                ms.Close();
                ms.Dispose();
                return true;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Email gönderim işlemi başarısız.{ex.InnerException}");

                return false;
            }

        }
    }
}
