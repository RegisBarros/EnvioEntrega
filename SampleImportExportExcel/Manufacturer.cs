using System;
using System.Collections.Generic;

namespace SampleImportExportExcel
{
    public class Manufacturer
    {
        public Guid Id { get; private set; }

        public string Name { get; set; }

        public string Industry { get; set; }

        public Manufacturer()
        {
            Id = Guid.NewGuid();
        }


        public static IEnumerable<Manufacturer> GetManufacturers()
        {
            return new List<Manufacturer>()
            {
                new Manufacturer()
                {
                    Name = "Hershey Co.",
                    Industry = "Food"
                },
                new Manufacturer()
                {
                    Name = "Apple Inc.",
                    Industry = "Computers & Other Electronic Products"
                },
                new Manufacturer()
                {
                    Name = "Western Refining Inc.",
                    Industry = "Petroleum & Coal Products"
                },
                new Manufacturer()
                {
                    Name = "Microsoft Corp.",
                    Industry = "Computers & Other Electronic Products"
                },
                new Manufacturer()
                {
                    Name = "Nike Inc.",
                    Industry = "Apparel"
                },
                new Manufacturer()
                {
                    Name = "IBM Corp.",
                    Industry = "Computers & Other Electronic Products"
                }
            };
        }
    }
}
