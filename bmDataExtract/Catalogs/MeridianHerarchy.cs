using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace bmDataExtract.Catalogs
{
    public class MeridianHerarchy
    {
        public Dictionary<int, Meridian> Meridians { get; set; }

        public Meridian SonOf(Meridian organMother)
        {
            if (organMother.ID == 2) return Meridians[12];
            if (organMother.ID == 5) return Meridians[8];
            foreach (var o in Meridians)
            {
                if (o.Value.Mother.ID == organMother.ID) return o.Value;
            }
            return null;
        }

        public Meridian SlaveOf(Meridian organMaster)
        {
            if (organMaster.ID == 2) return Meridians[1];
            if (organMaster.ID == 5) return Meridians[4];
            foreach (var o in Meridians)
            {
                if (o.Value.Master.ID == organMaster.ID) return o.Value;
            }
            return null;
        }

        public decimal IndGOf(Meridian organ)
        {
            decimal G = (organ.G_as_Master + organ.G_as_Mother + organ.Mother.G_as_Mother + organ.Master.G_as_Master) / 4;
            return G;
        }

        public MeridianHerarchy()
        {
            Meridians = new Dictionary<int, Meridian>();

            Meridians.Add(1,  new Meridian() { ID=1, Name = "Intestino Grueso", ShortName ="IG" });
            Meridians.Add(2,  new Meridian() { ID=2, Name = "Triple Recalentador", ShortName = "TR" });
            Meridians.Add(3,  new Meridian() { ID=3, Name = "Intestino Delgado", ShortName = "ID" });
            Meridians.Add(4,  new Meridian() { ID=4, Name = "Pulmon", ShortName = "P" });
            Meridians.Add(5,  new Meridian() { ID=5, Name = "Pericardio", ShortName = "PC" });
            Meridians.Add(6,  new Meridian() { ID=6, Name = "Corazon", ShortName = "C" });
            Meridians.Add(7,  new Meridian() { ID=7, Name = "Rinon", ShortName = "R" });
            Meridians.Add(8,  new Meridian() { ID=8, Name = "Bazo", ShortName = "B" });
            Meridians.Add(9,  new Meridian() { ID=9, Name = "Higado", ShortName = "H" });
            Meridians.Add(10, new Meridian() { ID=10,Name = "Vesicula Biliar", ShortName = "VB" });
            Meridians.Add(11, new Meridian() { ID=11,Name = "Vejiga", ShortName = "V" });
            Meridians.Add(12, new Meridian() { ID=12,Name = "Estomago", ShortName = "E" });

            #region Mother/Son relationship
            Meridians[1].Mother = Meridians[12];
            Meridians[2].Mother = Meridians[10];
            Meridians[3].Mother = Meridians[10];
            Meridians[4].Mother = Meridians[8];
            Meridians[5].Mother = Meridians[9];
            Meridians[6].Mother = Meridians[9];
            Meridians[7].Mother = Meridians[4];
            Meridians[8].Mother = Meridians[6];
            Meridians[9].Mother = Meridians[7];
            Meridians[10].Mother = Meridians[11];
            Meridians[11].Mother = Meridians[1];
            Meridians[12].Mother = Meridians[3];
            #endregion

            #region Master/Slave relationship
            Meridians[1].Master = Meridians[3];
            Meridians[2].Master = Meridians[11];
            Meridians[3].Master = Meridians[11];
            Meridians[4].Master = Meridians[6];
            Meridians[5].Master = Meridians[7];
            Meridians[6].Master = Meridians[7];
            Meridians[7].Master = Meridians[8];
            Meridians[8].Master = Meridians[9];
            Meridians[9].Master = Meridians[4];
            Meridians[10].Master = Meridians[1];
            Meridians[11].Master = Meridians[12];
            Meridians[12].Master = Meridians[10];
            #endregion
        }
    }
}
