using System;

namespace SampleImportExportExcel
{
    public class PropertyByName<T>
    {
        private object _propertyValue;

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

        public object PropertyValue
        {
            get
            {
                return _propertyValue;
            }
            set
            {
                _propertyValue = value;
            }
        }

        public int IntValue
        {
            get
            {
                if (PropertyValue == null || !int.TryParse(PropertyValue.ToString(), out int rez))
                    return default(int);
                return rez;
            }
        }

        public bool BooleanValue
        {
            get
            {
                if (PropertyValue == null || !bool.TryParse(PropertyValue.ToString(), out bool rez))
                    return default(bool);
                return rez;
            }
        }

        public string StringValue
        {
            get
            {
                return PropertyValue == null ? string.Empty : Convert.ToString(PropertyValue);
            }
        }

        public decimal DecimalValue
        {
            get
            {
                if (PropertyValue == null || !decimal.TryParse(PropertyValue.ToString(), out decimal rez))
                    return default(decimal);
                return rez;
            }
        }

        public decimal? DecimalValueNullable
        {
            get
            {
                if (PropertyValue == null || !decimal.TryParse(PropertyValue.ToString(), out decimal rez))
                    return null;
                return rez;
            }
        }

        public double DoubleValue
        {
            get
            {
                if (PropertyValue == null || !double.TryParse(PropertyValue.ToString(), out double rez))
                    return default(double);
                return rez;
            }
        }

        public DateTime? DateTimeNullable
        {
            get
            {
                if (PropertyValue == null || !DateTime.TryParse(StringValue, out DateTime rez))
                    return null;

                return rez;
            }
        }

        public Guid GuidValue
        {
            get
            {
                if (PropertyValue == null || !Guid.TryParse(StringValue, out Guid rez))
                    return default(Guid);

                return rez;
            }
        }

        public override string ToString()
        {
            return PropertyName;
        }
    }
}
