using Microsoft.Graph;
using System;
using System.Threading.Tasks;

namespace MicrosoftGraphEmailSender
{
    public class GraphEmailSender
    {
        private GraphServiceClient Client { get; set; }
        public GraphEmailSender(string clientId, string clientSecret, string tenantId)
        {
            var authProvider = new MicrosoftGraphAuthenticationProvider(clientId, clientSecret, new string[] { "https://graph.microsoft.com/.default" }, tenantId);
            Client = new GraphServiceClient(authProvider);
        }

        public async Task SendMailAsync(string from, string to, string message, bool isHTML = false)
        {
            ItemBody content = CreateItemBody(message, isHTML);
            Message msg = CreateMessage(from, to, content);
            await Client.Users[from].SendMail(msg, true).Request().PostAsync();
        }

        public void SendMail(string from, string to, string message, bool isHTML = false)
        {
            SendMailAsync(from, to, message, isHTML).Wait();
        }

        public void SendMail(string from, Message message)
        {
            SendMailAsync(from, message).Wait();
        }

        public async Task SendMailAsync(string from, Message message)
        {
            await Client.Users[from].SendMail(message, true).Request().PostAsync();
        }

        private Message CreateMessage(string from, string to, ItemBody content)
        {
            Message msg = new Message();
            msg.Body = content;
            msg.ToRecipients = new Recipient[] { new Recipient() { EmailAddress = new EmailAddress() { Address = to } } };
            return msg;

        }
        private ItemBody CreateItemBody(string content, bool isHtml = false)
        {
            ItemBody body = new ItemBody();
            body.ContentType = isHtml ? BodyType.Html : BodyType.Text;
            body.Content = content;
            return body;
        }
    }
}
