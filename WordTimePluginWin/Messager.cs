using System;
using System.Text;
using RabbitMQ.Client;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;

namespace WordTimePluginWin {
    internal class Messager {
        private IModel _channel;

        // 139.162.149.247 (wt_rabbitmq1)
        // 45.33.95.7 (wt_rabbitmq2)

        private const string _host = "139.162.149.247";
        private const int _port = 5672;
        private const string _user = "wordpluginwin";
        private const string _password = "worldig";

        public Messager() {
            var _factory = new ConnectionFactory() {
                VirtualHost = "wordtime",
                UserName = _user,
                Password = _password,
                Port = _port,
                HostName = _host
            };

            var connection = _factory.CreateConnection();
                _channel = connection.CreateModel();
            //queue: "wordtime",

            _channel.QueueDeclare(
                        durable: true,
                        exclusive: false,
                        autoDelete: false,
                        arguments: null);

                    var properties = _channel.CreateBasicProperties();
                    properties.Persistent = true;
                }

        struct Message {
            public string documentName;
            public string filePath;
            public DateTime messageCreationTime;
        };

        public string CreateMessage(ref Document document)
        {
            // var docProperties = document.BuiltInDocumentProperties;
            // var fileEditingTime = docProperties["Total editing time"];

            var message = new Message {
                documentName = document.Name,
                filePath = document.FullName,
                messageCreationTime = DateTime.UtcNow
            };

            var json = JsonConvert.SerializeObject(message);
            
            return json;
        }

        public void Send(string message) {
            var body = Encoding.UTF8.GetBytes(message);
                                
            _channel.BasicPublish(exchange: "fed.fanoutexchange",                 
                                    routingKey: "wordtime",
                                    basicProperties: null,
                                    body: body);
        }
    }
}