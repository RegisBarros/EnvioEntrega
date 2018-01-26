using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SampleImportExportExcel
{
    public  class ImportManager
    {
        public IEnumerable<Manufacturer> ImportManufacturersFromXlsx(Stream stream)
        {
            var manufacturers = new List<Manufacturer>();

            using (var xlPackage = new ExcelPackage(stream))
            {
                // get the first worksheet in the workbook
                var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                    throw new Exception("No worksheet found");

                //the columns
                var properties = GetPropertiesByExcelCells<Manufacturer>(worksheet);

                var manager = new PropertyManager<Manufacturer>(properties);

                var iRow = 2;

                while (true)
                {
                    var allColumnsAreEmpty = manager.GetProperties
                        .Select(property => worksheet.Cells[iRow, property.PropertyOrderPosition])
                        .All(cell => cell == null || cell.Value == null || string.IsNullOrEmpty(cell.Value.ToString()));

                    if (allColumnsAreEmpty)
                        break;

                    manager.ReadFromXlsx(worksheet, iRow);                    

                    var manufacturer = new Manufacturer();

                    foreach (var property in manager.GetProperties)
                    {
                        switch (property.PropertyName)
                        {
                            case "Id":
                                manufacturer.Id = property.GuidValue;
                                break;
                            case "Name":
                                manufacturer.Name = property.StringValue;
                                break;
                            case "Industry":
                                manufacturer.Industry = property.StringValue;
                                break;
                        }
                    }

                    manufacturers.Add(manufacturer);

                    iRow++;
                }
            }

            return manufacturers;
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
