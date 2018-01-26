using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;

namespace SampleImportExportExcel
{
    public class PropertyManager<T>
    {
        private readonly Dictionary<string, PropertyByName<T>> _properties;

        public T CurrentObject { get; set; }

        public PropertyManager(IEnumerable<PropertyByName<T>> properties)
        {
            _properties = new Dictionary<string, PropertyByName<T>>();

            var poz = 1;
            foreach (var propertyByName in properties)
            {
                propertyByName.PropertyOrderPosition = poz;
                poz++;
                _properties.Add(propertyByName.PropertyName, propertyByName);
            }
        }

        public void WriteCaption(ExcelWorksheet worksheet, Action<ExcelStyle> setStyle, int row = 1, int cellOffset = 0)
        {
            foreach (var caption in _properties.Values)
            {
                var cell = worksheet.Cells[row, caption.PropertyOrderPosition + cellOffset];
                cell.Value = caption;
                setStyle(cell.Style);
            }
        }

        public void WriteToXlsx(ExcelWorksheet worksheet, int row, int cellOffset = 0, ExcelWorksheet fWorksheet = null)
        {
            if (CurrentObject == null)
                return;

            foreach (var prop in _properties.Values)
            {
                var cell = worksheet.Cells[row, prop.PropertyOrderPosition + cellOffset];

                cell.Value = prop.GetProperty(CurrentObject);
            }
        }
    }
}
