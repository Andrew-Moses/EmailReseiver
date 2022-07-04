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
using System.Text.RegularExpressions;


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

        public static void CustomReplace(ImportData importData)
        {
           var obj_1 = Regex.Replace(importData.Quant.ToString(),"[.]",",") ;
           //var obj_2 = Regex.Replace(importData.Price.ToString(),"[.]",",") ;
           //var obj_3 = Regex.Replace(importData.PSum.ToString(),"[.]",",") ;


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

                        await client.Inbox.OpenAsync(MailKit.FolderAccess.ReadWrite);

                        var uids = await client.Inbox.SearchAsync(MailKit.Search.SearchQuery.NotSeen);

                        var messages = await client.Inbox.FetchAsync(uids,
                            MessageSummaryItems.Envelope | MessageSummaryItems.BodyStructure);

                        if (messages != null && messages.Count > 0)
                        {
                            foreach (var msg in messages)
                            {
                                client.Inbox.AddFlags(uids, MailKit.MessageFlags.Seen, true);

                                foreach (var att in msg.Attachments.OfType<BodyPartBasic>())
                                {
                                    var part = (MimePart)await client.Inbox.GetBodyPartAsync(msg.UniqueId, att);
                                    //if (!part.FileName.EndsWith("xlsx") || !part.FileName.EndsWith("xls")) continue;
                                    if (Regex.IsMatch(part.FileName, "XLSX") || Regex.IsMatch(part.FileName, "XLS")) continue;



                                    Stream outStream = new MemoryStream();
                                    await part.Content.DecodeToAsync(outStream);
                                    outStream.Position = 0;
                                    Spreadsheet document = new Spreadsheet();
                                    document.LoadFromStream(outStream);
                                    var sheet = document.Workbook.Worksheets[0];

                                    //client.Inbox.AddFlags(uids, MailKit.MessageFlags.Seen, { Silent = true});

                                    for (int row = 1; sheet.Cell(row, 0).ValueAsString != ""; row++)
                                    {
                                        ImportData importData = getData(sheet, row);
                                        //Console.WriteLine(importData.Quant);
                                        //CustomReplace(importData);
                                        // Запись в базу (dbo.ImportData)
                                        ImportData? _ = await _importDataService.AddEntry(importData);

                                    }
                                }
                            }
                        }
                        // TODO: где-то здесь удалить письмо из ящика или как-то запомнить uids, чтобы больше их не считывать
                        
                    }
                    //ожиадание полминуты до следующего цикла
                    await Task.Delay(30000);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private ImportData getData(Worksheet sheet, int row)
        {
            decimal quant = Convert.ToDecimal(sheet.Cell(row, 14).ValueAsString.Replace('.', ','));
            decimal price = Convert.ToDecimal(sheet.Cell(row, 16).ValueAsString.Replace('.', ','));
            decimal pSum = Convert.ToDecimal(sheet.Cell(row, 17).ValueAsString.Replace('.', ','));
            return new()
            {
                OrgName = sheet.Cell(row, 0).ValueAsString,
                MOD = sheet.Cell(row, 1).ValueAsString,
                INN = sheet.Cell(row, 2).ValueAsString,
                OKPO = sheet.Cell(row, 3).ValueAsString,
                FinancingItem = sheet.Cell(row, 4).ValueAsString,
                ProductName = sheet.Cell(row, 5).ValueAsString,
                MedForm = sheet.Cell(row, 6).ValueAsString,
                SeriaNum = sheet.Cell(row, 7).ValueAsString,
                MNN = sheet.Cell(row, 8).ValueAsString,
                MKB = sheet.Cell(row, 9).ValueAsString,
                RecSeria = sheet.Cell(row, 10).ValueAsString,
                RecNum = sheet.Cell(row, 11).ValueAsString,
                RecDate = sheet.Cell(row, 12).ValueAsDateTime,
                OtpuskDate = sheet.Cell(row, 13).ValueAsDateTime,
                Quant = quant,
                OkeiName = sheet.Cell(row, 15).ValueAsString,
                Price = price,
                PSum = pSum,
                LastName = sheet.Cell(row, 18).ValueAsString,
                Name = sheet.Cell(row, 19).ValueAsString,
                MidName = sheet.Cell(row, 20).ValueAsString,
                DateOB = sheet.Cell(row, 21).ValueAsDateTime,
                SNILS = sheet.Cell(row, 22).ValueAsString
            };
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