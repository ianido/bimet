using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace bmDataExtract.Catalogs
{
    public class Organ
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public Organ Mother { get; set; }
        public Organ Master { get; set; }
        public decimal G_as_Master { get; set; }
        public decimal G_as_Mother { get; set; }
        public decimal Variability { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
