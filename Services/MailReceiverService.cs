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
            services.AddScoped<ImportDataService>();
            var provider = services.BuildServiceProvider().CreateScope();
            _importDataService = provider.ServiceProvider.GetRequiredService<ImportDataService>();
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



                                    Stream outStream = new MemoryStream();
                                    await part.Content.DecodeToAsync(outStream);
                                    outStream.Position = 0;
                                    Spreadsheet document = new Spreadsheet();
                                    document.LoadFromStream(outStream);
                                    var sheet = document.Workbook.Worksheets[0];

                                    for (int row = 1; sheet.Cell(row, 0).ValueAsString != ""; row++)
                                    {

                                        ImportData importData = new ImportData
                                        {
                                            OrgName = sheet.Cell(row, 0).ValueAsString,
                                            MOD = sheet.Cell(row, 1).ValueAsString,
                                            ProductName = sheet.Cell(row, 2).ValueAsString,
                                            SeriaNum = sheet.Cell(row, 3).ValueAsString,
                                            MNN = sheet.Cell(row, 4).ValueAsString,
                                            RecNum = sheet.Cell(row, 5).ValueAsString,
                                            RecDate = sheet.Cell(row, 6).ValueAsDateTime,
                                            MedForm = sheet.Cell(row, 7).ValueAsString,
                                            Quant = (decimal)sheet.Cell(row, 8).ValueAsDouble,
                                            OkeiName = sheet.Cell(row, 9).ValueAsString,
                                            Price = (decimal)sheet.Cell(row, 10).ValueAsDouble,
                                            PSum = (decimal)sheet.Cell(row, 11).ValueAsDouble,
                                            LastName = sheet.Cell(row, 12).ValueAsString,
                                            Name = sheet.Cell(row, 13).ValueAsString,
                                            MidName = sheet.Cell(row, 14).ValueAsString,
                                            DateOB = sheet.Cell(row, 15).ValueAsDateTime,
                                            SNILS = sheet.Cell(row, 16).ValueAsInteger
                                        };


                                        // запись в базу
                                        ImportData? _ = await _importDataService.AddEntry(importData);

                                    }
                                }
                            }
                        }
                        // TODO: где-то здесь удалить почту из ящика или как-то запомнить uids, чтобы больше их не считывать

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

        private ImportDataService _importDataService;
    }
}