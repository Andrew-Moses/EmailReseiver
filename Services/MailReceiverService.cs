using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EmailReseiver.Contexts;
using EmailReseiver.Models;
using MailKit;
using MailKit.Search;
using Microsoft.Extensions.Configuration;
using MimeKit;
using MessageSummaryItems = MailKit.MessageSummaryItems;
using Bytescout.Spreadsheet;


namespace EmailReseiver.MailServices
{
    public class MailReceiverService
    {
        public IConfiguration Configuration { get; }

        public MailReceiverService()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(System.AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json",
                    optional: true,
                    reloadOnChange: true);
            Configuration = builder.Build();

            var connectionString = Configuration.GetConnectionString("LocalSql");

            IServiceCollection services = new ServiceCollection();
            services.AddDbContext<Context>(options => options.UseSqlServer(connectionString));
            services.AddScoped<DataBaseService>();
            var provider = services.BuildServiceProvider().CreateScope();
            _dataBaseService = provider.ServiceProvider.GetRequiredService<DataBaseService>();
        }

        public async Task DoReceiveMail()
        {
            var list = new List<MailItem>();
            var yandexUser = Configuration["YandexUser"];
            var yandexPass = Configuration["YandexPass"];
            try
            {
                while (true)
                {
                    using (var client = new MailKit.Net.Imap.ImapClient())
                    {
                        await client.ConnectAsync("imap.yandex.ru", 993, true);
                        await client.AuthenticateAsync(yandexUser, yandexPass);

                        await client.Inbox.OpenAsync(MailKit.FolderAccess.ReadOnly);

                        var uids = await client.Inbox.SearchAsync(SearchQuery.SentSince(DateTime.Now.AddDays(-7)));

                        var messages = await client.Inbox.FetchAsync(uids,
                            MessageSummaryItems.Envelope | MessageSummaryItems.BodyStructure);

                        if (messages != null && messages.Count > 0)
                        {
                            foreach (var msg in messages)
                            {
                                foreach (var att in msg.Attachments.OfType<BodyPartBasic>())
                                {
                                    var part = (MimePart)await client.Inbox.GetBodyPartAsync(msg.UniqueId, att);
                                    if (!part.FileName.EndsWith("xlsx")) continue;


                                    ImportData importData = new ();

                                    Stream outStream = new MemoryStream();
                                    await part.Content.DecodeToAsync(outStream);
                                    outStream.Position = 0;
                                    /*
                                                                        string fileName = String.Format(@"{0}.xlsx", System.Guid.NewGuid());
                                                                        await using var inputStream = File.Create(fileName);
                                                                        await part.Content.DecodeToAsync(inputStream);
                                                                        inputStream.Close();
                                                                        Spreadsheet document = new Spreadsheet();
                                                                        document.LoadFromFile(fileName);
                                                                        document.Workbook.Worksheets[0].SaveAsXML(outStream);
                                                                        outStream.Position = 0;
                                    */
                                    Spreadsheet document = new Spreadsheet();
                                    document.LoadFromStream(outStream);
                                    var sheet = document.Workbook.Worksheets[0];
                                    var rows = sheet.Rows;
                                    foreach (var row in rows)
                                    {

                                    }



                                    // запись в базу
                                    // await _dataBaseService.AddEntry(importData);
                                }
                            }
                        }
                    }

                    //ждем полминуты до следующего цикла
                    await Task.Delay(30000);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public static string StreamToString(Stream stream)
        {
            stream.Position = 0;
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private DataBaseService _dataBaseService;
    }
}