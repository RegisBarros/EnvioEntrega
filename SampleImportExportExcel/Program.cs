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
            //var configuration = new ConfigurationBuilder()
            //          .SetBasePath(Directory.GetCurrentDirectory())
            //         .AddJsonFile("appsettings.json")
            //         .Build();

            string directory = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(directory, "Files/entregas.xlsx");

            Stream file = File.OpenRead(filePath);

            var importManager = new ImportManager();
            IEnumerable<Order> orders = importManager.ImportOrdersFromXlsx(file);


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
