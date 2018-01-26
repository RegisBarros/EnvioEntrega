using Microsoft.Extensions.Configuration;
using Saraiva.Framework.Service.Email;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SampleImportExportExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var configuration = new ConfigurationBuilder()
                      .SetBasePath(Directory.GetCurrentDirectory())
                     .AddJsonFile("appsettings.json")
                     .Build();

            var exportManger = new ExportManager();
            var bytes = exportManger.ExportManufacturersToXlsx(Manufacturer.GetManufacturers());

            SendEmail(bytes);

            Console.ReadKey();
        }

        static void SendEmail(byte[] file)
        {
            var servico = EmailService.Instance;

            IList<EnvioArquivoStreamDTO> atachements = new List<EnvioArquivoStreamDTO>()
            {
                new EnvioArquivoStreamDTO()
                {
                    Conteudo = new MemoryStream(file.ToArray()),
                    Formato = "xlsx",
                    NomeArquivo = "PlanilhaTeste"
                }
            };

            string[] to = { "reginaldo.barros@fcamara.com.br" }; 

            servico.EnviarEmail(to, "SheetTest", "test file xlsx", atachements, true).Wait();
        }
    }
}
