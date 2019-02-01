using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace bmDataExtract.Catalogs
{
    public class OrgansHerarchy
    {
        public Dictionary<int, Organ> Organs { get; set; }

        public Organ SonOf(Organ organMother)
        {
            foreach (var o in Organs)
            {
                if (o.Value.Mother.ID == organMother.ID) return o.Value;
            }
            return null;
        }

        public Organ SlaveOf(Organ organMaster)
        {
            foreach (var o in Organs)
            {
                if (o.Value.Master.ID == organMaster.ID) return o.Value;
            }
            return null;
        }

        public decimal IndGOf(Organ organ)
        {
            decimal G = (organ.G_as_Master + organ.G_as_Mother + organ.Mother.G_as_Mother + organ.Master.G_as_Master) / 4;
            return G;
        }

        public OrgansHerarchy()
        {
            Organs = new Dictionary<int, Organ>();

            Organs.Add(1,  new Organ() { ID=1, Name = "Intestino Grueso" });
            Organs.Add(2,  new Organ() { ID=2, Name = "Triple Recalentador" });
            Organs.Add(3,  new Organ() { ID=3, Name = "Intestino Delgado" });
            Organs.Add(4,  new Organ() { ID=4, Name = "Pulmon" });
            Organs.Add(5,  new Organ() { ID=5, Name = "Pericardio" });
            Organs.Add(6,  new Organ() { ID=6, Name = "Corazon" });
            Organs.Add(7,  new Organ() { ID=7, Name = "Rinon" });
            Organs.Add(8,  new Organ() { ID=8, Name = "Bazo" });
            Organs.Add(9,  new Organ() { ID=9, Name = "Higado" });
            Organs.Add(10, new Organ() { ID=10,Name = "Vesicula Biliar" });
            Organs.Add(11, new Organ() { ID=11,Name = "Vejiga" });
            Organs.Add(12, new Organ() { ID=12,Name = "Estomago" });

            #region Mother/Son relationship
            Organs[1].Mother = Organs[12];
            Organs[2].Mother = Organs[10];
            Organs[3].Mother = Organs[10];
            Organs[4].Mother = Organs[8];
            Organs[5].Mother = Organs[9];
            Organs[6].Mother = Organs[9];
            Organs[7].Mother = Organs[4];
            Organs[8].Mother = Organs[6];
            Organs[9].Mother = Organs[7];
            Organs[10].Mother = Organs[11];
            Organs[11].Mother = Organs[10];
            Organs[12].Mother = Organs[3];
            #endregion

            #region Master/Slave relationship
            Organs[1].Master = Organs[3];
            Organs[2].Master = Organs[11];
            Organs[3].Master = Organs[11];
            Organs[4].Master = Organs[6];
            Organs[5].Master = Organs[7];
            Organs[6].Master = Organs[7];
            Organs[7].Master = Organs[8];
            Organs[8].Master = Organs[9];
            Organs[9].Master = Organs[4];
            Organs[10].Master = Organs[1];
            Organs[11].Master = Organs[12];
            Organs[12].Master = Organs[10];
            #endregion
        }
    }
}
