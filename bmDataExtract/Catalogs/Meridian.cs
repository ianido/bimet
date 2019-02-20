using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace bmDataExtract.Catalogs
{
    public class Meridian
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public Meridian Mother { get; set; }
        public Meridian Master { get; set; }
        public decimal I_Bioene { get; set; }
        public decimal LeftPotential { get; set; }
        public decimal RightPotential { get; set; }
        public decimal G_as_Master { get; set; }
        public decimal G_as_Mother { get; set; }
        public decimal Variability { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
