using bmDataExtract.Catalogs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace bmDataExtract
{
    public class ExtractBimet
    {
        public string Dir { get; private set; }
        public string DB { get; private set; }
        public string[] PatientsID { get; private set; }

        private OrgansHerarchy str = new OrgansHerarchy();
        public ExtractBimet(string directory, string db)
        {
            Dir = directory;
            DB = db;
            string[] _tFiles = Directory.GetFiles(Dir, "*ATOT.1");
            List<string> fList = new List<string>();
            foreach (string f in _tFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(f);
                fList.Add(fileName.Replace("ATOT", ""));
            }

            PatientsID = fList.ToArray();
        }

        public void Start()
        {
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(DB)))
            {
                var sheet = excel.Workbook.Worksheets[3];
                // RSearch Code from Range H6:H..
                int startRow = 6;
                string patientCode = sheet.Cells["H" + startRow].Text;

                while (!string.IsNullOrEmpty(patientCode))
                {
                    try
                    {
                        Console.WriteLine("Processing Patient: "+ patientCode);
                        Step1(sheet, patientCode, startRow);
                        Step2(sheet, patientCode, startRow);
                        Step3(sheet, patientCode, startRow);
                        Step4(sheet, patientCode, startRow);
                    }
                    catch (FileNotFoundException)
                    {
                        Console.WriteLine($"Patient DB for: {patientCode} not found.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error Processing Patient: {patientCode}." + ex.Message);
                    }
                    startRow++;
                    patientCode = sheet.Cells["H" + startRow].Text;
                }

                excel.SaveAs(new FileInfo(DB));
            }
        }


        /// <summary>
        /// Damaged Organs, IND G and % V extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step1(ExcelWorksheet sheet, string patientCode, int row)
        {

            BimetOneReader bor = new BimetOneReader(Path.Combine(Dir, $"{patientCode}ind.1"), false);

            // Set Organs Name
            sheet.Cells[$"I{row}"].Value = str.Organs[Convert.ToInt32(bor[0])];
            sheet.Cells[$"L{row}"].Value = str.Organs[Convert.ToInt32(bor[1])];
            sheet.Cells[$"O{row}"].Value = str.Organs[Convert.ToInt32(bor[2])];

            // Set IND G
            sheet.Cells[$"K{row}"].Value = Convert.ToDecimal(bor[8]);
            sheet.Cells[$"N{row}"].Value = Convert.ToDecimal(bor[9]);
            sheet.Cells[$"Q{row}"].Value = Convert.ToDecimal(bor[10]);

            // Set IND G
            sheet.Cells[$"J{row}"].Value = Convert.ToDecimal(bor[35]);
            sheet.Cells[$"M{row}"].Value = Convert.ToDecimal(bor[36]);
            sheet.Cells[$"P{row}"].Value = Convert.ToDecimal(bor[37]);

        }

        /// <summary>
        /// Variabilidad Percentage Extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step2(ExcelWorksheet sheet, string patientCode, int row)
        {
            BimetOneReader bor = new BimetOneReader(Path.Combine(Dir, $"{patientCode}vari.1"), true);
            int lastSession = bor.Sessions[bor.Sessions.Length - 1];
            foreach (var org in str.Organs)
            {
                org.Value.Variability = Convert.ToDecimal(bor[lastSession, org.Key-1]);
            }


            // Set Variability % for Corazon
            sheet.Cells[$"V{row}"].Value = str.Organs[6].Variability;
            // Set Variability % for Intestino Delgado
            sheet.Cells[$"AA{row}"].Value = str.Organs[3].Variability;
            // Set Variability % for Triple R
            sheet.Cells[$"AF{row}"].Value = str.Organs[2].Variability;
            // Set Variability % for Pericardio
            sheet.Cells[$"AK{row}"].Value = str.Organs[5].Variability;
            // Set Variability % for Bazo
            sheet.Cells[$"AP{row}"].Value = str.Organs[8].Variability;
            // Set Variability % for Estomago
            sheet.Cells[$"AU{row}"].Value = str.Organs[12].Variability;
            // Set Variability % for Pulmon
            sheet.Cells[$"AZ{row}"].Value = str.Organs[4].Variability;
            // Set Variability % for Intestion Grueso
            sheet.Cells[$"BE{row}"].Value = str.Organs[1].Variability;
            // Set Variability % for Rinon
            sheet.Cells[$"BJ{row}"].Value = str.Organs[7].Variability;
            // Set Variability % for Vejiga
            sheet.Cells[$"BO{row}"].Value = str.Organs[11].Variability;
            // Set Variability % for Higado
            sheet.Cells[$"BT{row}"].Value = str.Organs[9].Variability;
            // Set Variability % for Vesicula Biliar
            sheet.Cells[$"BY{row}"].Value = str.Organs[10].Variability;
        }

        /// <summary>
        /// Madre y Master extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step3(ExcelWorksheet sheet, string patientCode, int row)
        {
            BimetOneReader bor = new BimetOneReader(Path.Combine(Dir, $"{patientCode}ener.1"), true);
            int lastSession = bor.Sessions[bor.Sessions.Length - 1];
            foreach (var org in str.Organs)
            {
                org.Value.G_as_Mother = Convert.ToDecimal(bor[lastSession, org.Key-1]);
                org.Value.G_as_Master = Convert.ToDecimal(bor[lastSession, org.Key+11])*-1;
            }

            decimal maxLimit = 0.3M;
            
            // Set for Corazon
            if (str.Organs[6].G_as_Mother > maxLimit) sheet.Cells[$"S{row}"].Value = str.Organs[6].G_as_Mother;
            if (str.Organs[6].G_as_Master > maxLimit) sheet.Cells[$"T{row}"].Value = str.Organs[6].G_as_Master;
            // Set for Intestino Delgado
            if (str.Organs[3].G_as_Mother > maxLimit) sheet.Cells[$"X{row}"].Value = str.Organs[3].G_as_Mother;
            if (str.Organs[3].G_as_Master > maxLimit) sheet.Cells[$"Y{row}"].Value = str.Organs[3].G_as_Master;
            // Set for Triple R
            if (str.Organs[2].G_as_Mother > maxLimit) sheet.Cells[$"AC{row}"].Value = str.Organs[2].G_as_Mother;
            if (str.Organs[2].G_as_Master > maxLimit) sheet.Cells[$"AD{row}"].Value = str.Organs[2].G_as_Master;
            // Set for Pericardio
            if (str.Organs[5].G_as_Mother > maxLimit) sheet.Cells[$"AH{row}"].Value = str.Organs[5].G_as_Mother;
            if (str.Organs[5].G_as_Master > maxLimit) sheet.Cells[$"AI{row}"].Value = str.Organs[5].G_as_Master;
            // Set for Bazo
            if (str.Organs[8].G_as_Mother > maxLimit) sheet.Cells[$"AM{row}"].Value = str.Organs[8].G_as_Mother;
            if (str.Organs[8].G_as_Master > maxLimit) sheet.Cells[$"AN{row}"].Value = str.Organs[8].G_as_Master;
            // Set for Estomago
            if (str.Organs[12].G_as_Mother > maxLimit) sheet.Cells[$"AR{row}"].Value = str.Organs[12].G_as_Mother;
            if (str.Organs[12].G_as_Master > maxLimit) sheet.Cells[$"AS{row}"].Value = str.Organs[12].G_as_Master;
            // Set for Pulmon
            if (str.Organs[4].G_as_Mother > maxLimit) sheet.Cells[$"AW{row}"].Value = str.Organs[4].G_as_Mother;
            if (str.Organs[4].G_as_Master > maxLimit) sheet.Cells[$"AX{row}"].Value = str.Organs[4].G_as_Master;
            // Set for Intestion Grueso
            if (str.Organs[1].G_as_Mother > maxLimit) sheet.Cells[$"BB{row}"].Value = str.Organs[1].G_as_Mother;
            if (str.Organs[1].G_as_Master > maxLimit) sheet.Cells[$"BC{row}"].Value = str.Organs[1].G_as_Master;
            // Set for Rinon
            if (str.Organs[7].G_as_Mother > maxLimit) sheet.Cells[$"BG{row}"].Value = str.Organs[7].G_as_Mother;
            if (str.Organs[7].G_as_Master > maxLimit) sheet.Cells[$"BH{row}"].Value = str.Organs[7].G_as_Master;
            // Set for Vejiga
            if (str.Organs[11].G_as_Mother > maxLimit) sheet.Cells[$"BL{row}"].Value = str.Organs[11].G_as_Mother;
            if (str.Organs[11].G_as_Master > maxLimit) sheet.Cells[$"BM{row}"].Value = str.Organs[11].G_as_Master;
            // Set for Higado
            if (str.Organs[9].G_as_Mother > maxLimit) sheet.Cells[$"BQ{row}"].Value = str.Organs[9].G_as_Mother;
            if (str.Organs[9].G_as_Master > maxLimit) sheet.Cells[$"BR{row}"].Value = str.Organs[9].G_as_Master;
            // Set for Vesicula Biliar
            if (str.Organs[10].G_as_Mother > maxLimit) sheet.Cells[$"BV{row}"].Value = str.Organs[10].G_as_Mother;
            if (str.Organs[10].G_as_Master > maxLimit) sheet.Cells[$"BW{row}"].Value = str.Organs[10].G_as_Master;
        }

        /// <summary>
        /// Ing G extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step4(ExcelWorksheet sheet, string patientCode, int row)
        {
            sheet.Cells[$"U{row}"].Value  = str.IndGOf(str.Organs[6]);
            sheet.Cells[$"Z{row}"].Value  = str.IndGOf(str.Organs[3]);
            sheet.Cells[$"AE{row}"].Value = str.IndGOf(str.Organs[2]);
            sheet.Cells[$"AJ{row}"].Value = str.IndGOf(str.Organs[5]);
            sheet.Cells[$"AO{row}"].Value = str.IndGOf(str.Organs[8]);
            sheet.Cells[$"AT{row}"].Value = str.IndGOf(str.Organs[12]);
            sheet.Cells[$"AY{row}"].Value = str.IndGOf(str.Organs[4]);
            sheet.Cells[$"BD{row}"].Value = str.IndGOf(str.Organs[1]);
            sheet.Cells[$"BI{row}"].Value = str.IndGOf(str.Organs[7]);
            sheet.Cells[$"BN{row}"].Value = str.IndGOf(str.Organs[11]);
            sheet.Cells[$"BS{row}"].Value = str.IndGOf(str.Organs[9]);
            sheet.Cells[$"BX{row}"].Value = str.IndGOf(str.Organs[10]);
        }
    }
}
