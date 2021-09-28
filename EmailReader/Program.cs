using System;
using Microsoft.Exchange.WebServices.Data;

namespace EmailReader
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService _service;
            int itemsToRetrieve = 5;
            try
            {
                Console.WriteLine("Registering Exchange connection");

                var creds = new Credentials();
                _service = new ExchangeService
                {
                    Credentials = new WebCredentials(creds.UserName, creds.Password)
                };
            }
            catch
            {
                Console.WriteLine("Connection to Exchange failed.");
                return;
            }

            _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            try
            {
                Console.WriteLine("Reading mail");

                var findResults = _service.FindItems(WellKnownFolderName.Inbox, new ItemView(itemsToRetrieve));
                foreach (Item item in findResults.Items)
                {
                    EmailMessage message = EmailMessage.Bind(_service, item.Id, new PropertySet(ItemSchema.Attachments));

                    message.Load();
                    Console.WriteLine("=============================================================");
                    Console.WriteLine($"Message from {message.From}. Attachment Count: {message.Attachments.Count}");
                    //Console.WriteLine($"{message.Body.Text}");
                }

                Console.WriteLine("Done.");
            }
            catch (Exception e)
            {
                Console.WriteLine("An error has occured. \n:" + e.Message);
            }
        }
    }
}
