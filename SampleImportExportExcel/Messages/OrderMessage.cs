using System;

namespace SampleImportExportExcel.Messages
{
    public class OrderMessage
    {
        public string Numero { get; set; }

        public Rastreio Rastreio { get; set; }

        public NotaFiscal NotaFiscal { get; set; }

        public static OrderMessage Create(Order order)
        {
            return new OrderMessage
            {
                Numero = order.Pedido,
                NotaFiscal = new NotaFiscal
                {
                    ChaveAcesso = "JN344225297BR",
                    Serie = "1010",
                    Numero = "102030",
                    CaminhoUrl = "http://saraiva.com.br",
                    Data = DateTime.Now,
                    Skus = new string[] { order.SkuLojista }
                },
                Rastreio = new Rastreio
                {
                    URL = "https://correiosrastrear.com",
                    Numero = order.Rastreio,
                    Transportadora = "Saraiva",
                    DataTracking = DateTime.Now,
                    SkusLojista = new string[] { order.SkuLojista }
                }
            };
        }
    }

    public class Rastreio
    {
        public string URL { get; set; }

        public string Numero { get; set; }
        public string Transportadora { get; set; }
        public DateTime DataTracking { get; set; }
        public string[] SkusLojista { get; set; }
    }

    public class NotaFiscal
    {
        public string ChaveAcesso { get; set; }
        public string Serie { get; set; }
        public string Numero { get; set; }
        public string CaminhoUrl { get; set; }
        public DateTime Data { get; set; }
        public string[] Skus { get; set; }
    }
}

/*{
  "Numero": "JN344225297BR",
  "Rastreio": {
    "URL": "https://correiosrastrear.com",
    "Numero": "JN344225297BR",
    "Transportadora": "Saraiva",
    "DataTracking": "2019-06-19 19:00:28",
    "SkusLojista": [
      "182587"
    ]
  },
  "NotaFiscal": {
    "ChaveAcesso": "JN344225297BR",
    "Serie": "1010",
    "Numero": "102030",
    "CaminhoUrl": "http://saraiva.com.br",
    "Data": "2019-06-19 19:00:28",
    "Skus": [
      "182587"
    ]
  }
} */
