using System;
using System.Text;
using RabbitMQ.Client;

namespace WordTimePluginWin {
    internal class Messager {
        private IModel _channel;

        private const string _host = "139.162.149.247";
        private const int _port = 5672;
        private const string _user = "wordpluginwin";
        private const string _password = "worldig";

        public Messager() {
            var _factory = new ConnectionFactory() {
                VirtualHost = "/",
                UserName = _user,
                Password = _password,
                Port = _port,
                HostName = _host
            };

            var connection = _factory.CreateConnection();
                _channel = connection.CreateModel();

                    _channel.QueueDeclare(queue: "wordtime",
                        durable: true,
                        exclusive: false,
                        autoDelete: false,
                        arguments: null);

                    var properties = _channel.CreateBasicProperties();
                    properties.Persistent = true;
                }
            
        

        public void Send(string message) {
            var body = Encoding.UTF8.GetBytes(message);
                    
            _channel.BasicPublish(exchange: "", 
                
                                    routingKey: "wordtime",
                                    basicProperties: null,
                                    body: body);;
        }
    }
}