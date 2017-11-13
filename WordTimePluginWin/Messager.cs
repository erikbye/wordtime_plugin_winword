using System;
using System.Text;
using RabbitMQ.Client;

namespace WordTimePluginWin {
    internal class Messager {

        private static readonly ConnectionFactory _factory;
        private const string _host = "45.79.139.68";
        private const int _port = 5672;
        private const string _user = "wordpluginwin";
        private const string _password = "worldig";

        static Messager() {
            _factory = new ConnectionFactory() {               
                VirtualHost = "/",
                UserName = _user,
                Password = _password,
                Port = _port,
                HostName = _host                
            };
        }

        public static void Send(string message) {

            using (var connection = _factory.CreateConnection()) {
                using (var channel = connection.CreateModel()) {
                    channel.QueueDeclare(queue: "hello",
                                         durable: false,
                                         exclusive: false,
                                         autoDelete: false,
                                         arguments: null);

                    var properties = channel.CreateBasicProperties();                    
                    properties.Persistent = true;

                    var body = Encoding.UTF8.GetBytes(message);
                    
                    channel.BasicPublish(exchange: "",
                                         routingKey: "hello",
                                         basicProperties: null,
                                         body: body);;
                }
            }
        }
    }
}