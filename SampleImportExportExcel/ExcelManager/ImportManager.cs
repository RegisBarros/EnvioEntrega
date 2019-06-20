using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SampleImportExportExcel
{
    public  class ImportManager
    {
        public IEnumerable<Order> ImportOrdersFromXlsx(Stream stream)
        {
            var orders = new List<Order>();

            using (var xlPackage = new ExcelPackage(stream))
            {
                // get the first worksheet in the workbook
                var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                    throw new Exception("No worksheet found");

                //the columns
                var properties = GetPropertiesByExcelCells<Order>(worksheet);

                var manager = new PropertyManager<Order>(properties);

                var iRow = 2;

                while (true)
                {
                    var allColumnsAreEmpty = manager.GetProperties
                        .Select(property => worksheet.Cells[iRow, property.PropertyOrderPosition])
                        .All(cell => cell == null || cell.Value == null || string.IsNullOrEmpty(cell.Value.ToString()));

                    if (allColumnsAreEmpty)
                        break;

                    manager.ReadFromXlsx(worksheet, iRow);

                    var order = new Order();

                    foreach (var property in manager.GetProperties)
                    {
                        switch (property.PropertyName)
                        {
                            case "Pedido":
                                order.Pedido = property.StringValue;
                                break;
                            case "SkuLojista":
                                order.SkuLojista = property.StringValue;
                                break;
                            case "Rastreio":
                                order.Rastreio = property.StringValue;
                                break;
                        }
                    }

                    orders.Add(order);

                    iRow++;
                }
            }

            return orders;
        }

        public static IList<PropertyByName<T>> GetPropertiesByExcelCells<T>(ExcelWorksheet worksheet)
        {
            var properties = new List<PropertyByName<T>>();
            var poz = 1;
            while (true)
            {
                try
                {
                    var cell = worksheet.Cells[1, poz];

                    if (string.IsNullOrEmpty(cell?.Value?.ToString()))
                        break;

                    poz += 1;
                    properties.Add(new PropertyByName<T>(cell.Value.ToString()));
                }
                catch
                {
                    break;
                }
            }

            return properties;
        }
    }
}
