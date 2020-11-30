using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordToPdfConverter.Producer.Models;

namespace WordToPdfConverter.Producer.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;

        public HomeController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult WordToPdfPage()
        {

            return View();
            
        }
        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdf wordToPdf)
        {
            var factory = new ConnectionFactory();
            factory.Uri = new Uri(_configuration["ConnectionStrings:RabbitMQCloudString"]);

            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {

                    channel.ExchangeDeclare("convertExchange", ExchangeType.Direct, true, false, null);

                    channel.QueueDeclare(queue: "File", durable: true, exclusive: false, autoDelete: false, arguments: null);

                    channel.QueueBind(queue: "File", exchange: "convertExchange", routingKey: "WordToPdf");

                    MessagWordToPdf message = new MessagWordToPdf();

                    using (MemoryStream ms = new MemoryStream())
                    {
                        wordToPdf.WordFile.CopyTo(ms);
                        message.WordByte = ms.ToArray();
                    }

                    message.Email = wordToPdf.Email;
                    message.FileName = Path.GetFileNameWithoutExtension(wordToPdf.WordFile.FileName);

                    string serializeMessage = JsonConvert.SerializeObject(message);
                    byte[] byteMessage = Encoding.UTF8.GetBytes(serializeMessage);

                    var properties = channel.CreateBasicProperties();
                    properties.Persistent = true;

                    channel.BasicPublish(exchange: "convertExchange", routingKey: "WordToPdf", basicProperties: properties,body:byteMessage);
                   
                    ViewBag.result = "Dosyanız işleme alınmıştır. İşlem bittiğinde dosyanız email ile gönderilecektir.";

                    return View();
                }
            }
        }
        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
