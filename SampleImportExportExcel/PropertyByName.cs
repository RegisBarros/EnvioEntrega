using System;
using System.Collections.Generic;
using System.Text;

namespace SampleImportExportExcel
{
    public class PropertyByName<T>
    {
        public int PropertyOrderPosition { get; set; }

        public Func<T, object> GetProperty { get; }

        public string PropertyName { get; }

        public bool Ignore { get; set; }

        public PropertyByName(string propertyName, Func<T, object> func = null, bool ignore = false)
        {
            this.PropertyName = propertyName;
            this.GetProperty = func;
            this.PropertyOrderPosition = 1;
            this.Ignore = ignore;
        }

        public override string ToString()
        {
            return PropertyName;
        }
    }
}
