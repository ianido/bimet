﻿using bmDataExtract.Catalogs;
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
        private ILogger logger { get; set; }
        public int HeadRow { get; set; }



        public enum CategoryColor
        {
            Dominancia,
            Contracorriente,
            Bloqueo,
            Contradominancia        
        }

        private MeridianHerarchy str = new MeridianHerarchy();

        private void SetCategoryColor(ExcelRange excelRange, CategoryColor category)
        {
            excelRange.Style.Font.Color.SetColor(1, 0, 0, 0);
            excelRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            switch (category)
            {
                case CategoryColor.Contracorriente:
                    excelRange.Style.Fill.BackgroundColor.SetColor(1, 240, 90, 90);
                    break;
                case CategoryColor.Contradominancia:
                    excelRange.Style.Fill.BackgroundColor.SetColor(1, 220, 90, 240);
                    break;
                case CategoryColor.Bloqueo:
                    excelRange.Style.Fill.BackgroundColor.SetColor(1, 90, 200, 240);
                    break;
                case CategoryColor.Dominancia:
                    excelRange.Style.Fill.BackgroundColor.SetColor(1, 90, 240, 170);
                    break;
            }
        }

        // (1 = A, 2 = B...27 = AA...703 = AAA...)
        public static string GetColNameFromIndex(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        // (A = 1, B = 2...AA = 27...AAA = 703...)
        public static int GetColNumberFromName(string columnName)
        {
            char[] characters = columnName.ToUpperInvariant().ToCharArray();
            int sum = 0;
            for (int i = 0; i < characters.Length; i++)
            {
                sum *= 26;
                sum += (characters[i] - 'A' + 1);
            }
            return sum;
        }

        public ExtractBimet(string directory, string db, ILogger ilogger)
        {
            Dir = directory;
            DB = db;
            logger = ilogger;
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
            logger.Log("Start Processing: " + DB + "...", EventType.Info);
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(DB)))
            {
                var sheet = excel.Workbook.Worksheets[3];
                // RSearch Code from Range H6:H..
                int startRow = 6;
                HeadRow = 5;
                string patientCode = sheet.Cells["H" + startRow].Text;

                while (!string.IsNullOrEmpty(patientCode))
                {
                    try
                    {
                        logger.Log("Processing Patient: "+ patientCode, finishLine : false);
                        Step2(sheet, patientCode, startRow);
                        Step3(sheet, patientCode, startRow);
                        Step3a(sheet, patientCode, startRow);
                        Step4(sheet, patientCode, startRow);
                        Step5(sheet, patientCode, startRow);
                        logger.Log($" --> OK.", EventType.Success);
                    }
                    catch (FileNotFoundException)
                    {
                        logger.Log($" --> Not found.", EventType.Warning);
                    }
                    catch (Exception ex)
                    {
                        logger.Log($" --> Error:" + ex.Message, EventType.Error);
                    }
                    startRow++;
                    patientCode = sheet.Cells["H" + startRow].Text;
                }
                logger.Log("");
                logger.Log("Saving DB (Excel) changes...", EventType.Info);
                excel.SaveAs(new FileInfo(DB));
            }
        }
        
        /// <summary>
        /// Damaged Organs, IND G and % V extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step3a(ExcelWorksheet sheet, string patientCode, int row)
        {
            str.Meridians[6].G_Total = str.IndGOf(str.Meridians[6]);
            str.Meridians[3].G_Total = str.IndGOf(str.Meridians[3]);
            str.Meridians[2].G_Total = str.IndGOf(str.Meridians[2]);
            str.Meridians[5].G_Total = str.IndGOf(str.Meridians[5]);
            str.Meridians[8].G_Total = str.IndGOf(str.Meridians[8]);
            str.Meridians[12].G_Total = str.IndGOf(str.Meridians[12]);
            str.Meridians[4].G_Total = str.IndGOf(str.Meridians[4]);
            str.Meridians[1].G_Total = str.IndGOf(str.Meridians[1]);
            str.Meridians[7].G_Total = str.IndGOf(str.Meridians[7]);
            str.Meridians[11].G_Total = str.IndGOf(str.Meridians[11]);
            str.Meridians[9].G_Total = str.IndGOf(str.Meridians[9]);
            str.Meridians[10].G_Total = str.IndGOf(str.Meridians[10]);

            var ol = str.Meridians.OrderByDescending(e => e.Value.G_Total).ToArray();

            // Set Organs Name
            sheet.Cells[$"I{row}"].Value = ol[0].Value.Name;
            sheet.Cells[$"L{row}"].Value = ol[1].Value.Name;
            sheet.Cells[$"O{row}"].Value = ol[2].Value.Name;

            // Set Variab
            sheet.Cells[$"K{row}"].Value = ol[0].Value.Variability;
            sheet.Cells[$"N{row}"].Value = ol[1].Value.Variability;
            sheet.Cells[$"Q{row}"].Value = ol[2].Value.Variability;

            // Set IND G
            sheet.Cells[$"J{row}"].Value = ol[0].Value.G_Total;
            sheet.Cells[$"M{row}"].Value = ol[1].Value.G_Total;
            sheet.Cells[$"P{row}"].Value = ol[2].Value.G_Total;

        }

        /// <summary>
        /// Variabilidad Percentage Extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step2(ExcelWorksheet sheet, string patientCode, int row)
        {
            BimetOneReader bor = new BimetOneReader(Path.Combine(Dir, $"{patientCode}vari.1"), BimetOneReader.FileFormat.NumeratedSessions);
            int lastSession = bor.Sessions[bor.Sessions.Length - 1];
            foreach (var org in str.Meridians)
            {
                org.Value.Variability = Convert.ToDecimal(bor[lastSession, org.Key-1]);
            }


            // Set Variability % for Corazon
            sheet.Cells[$"V{row}"].Value = str.Meridians[6].Variability;
            // Set Variability % for Intestino Delgado
            sheet.Cells[$"AA{row}"].Value = str.Meridians[3].Variability;
            // Set Variability % for Triple R
            sheet.Cells[$"AF{row}"].Value = str.Meridians[2].Variability;
            // Set Variability % for Pericardio
            sheet.Cells[$"AK{row}"].Value = str.Meridians[5].Variability;
            // Set Variability % for Bazo
            sheet.Cells[$"AP{row}"].Value = str.Meridians[8].Variability;
            // Set Variability % for Estomago
            sheet.Cells[$"AU{row}"].Value = str.Meridians[12].Variability;
            // Set Variability % for Pulmon
            sheet.Cells[$"AZ{row}"].Value = str.Meridians[4].Variability;
            // Set Variability % for Intestion Grueso
            sheet.Cells[$"BE{row}"].Value = str.Meridians[1].Variability;
            // Set Variability % for Rinon
            sheet.Cells[$"BJ{row}"].Value = str.Meridians[7].Variability;
            // Set Variability % for Vejiga
            sheet.Cells[$"BO{row}"].Value = str.Meridians[11].Variability;
            // Set Variability % for Higado
            sheet.Cells[$"BT{row}"].Value = str.Meridians[9].Variability;
            // Set Variability % for Vesicula Biliar
            sheet.Cells[$"BY{row}"].Value = str.Meridians[10].Variability;
        }

        /// <summary>
        /// Madre y Master extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step3(ExcelWorksheet sheet, string patientCode, int row)
        {
            BimetOneReader bor = new BimetOneReader(Path.Combine(Dir, $"{patientCode}ener.1"), BimetOneReader.FileFormat.NumeratedSessions);
            int lastSession = bor.Sessions[bor.Sessions.Length - 1];
            foreach (var org in str.Meridians)
            {
                org.Value.G_as_Mother = Convert.ToDecimal(bor[lastSession, org.Key-1]);
                org.Value.G_as_Master = Convert.ToDecimal(bor[lastSession, org.Key+11])*-1;
            }

            BimetOneReader col = new BimetOneReader(Path.Combine(Dir, $"{patientCode}.1"), BimetOneReader.FileFormat.UnnumeratedSessions);
            lastSession = col.Sessions[col.Sessions.Length - 1];
            foreach (var org in str.Meridians)
            {
                org.Value.I_Bioene = Convert.ToDecimal(col[lastSession, org.Key - 1]);                
            }

            decimal maxLimit = 0.3M;

            sheet.Cells[$"S{row}"].Value = str.Meridians[6].G_as_Mother;
            sheet.Cells[$"T{row}"].Value = str.Meridians[6].G_as_Master;
            sheet.Cells[$"X{row}"].Value = str.Meridians[3].G_as_Mother;
            sheet.Cells[$"Y{row}"].Value = str.Meridians[3].G_as_Master;
            sheet.Cells[$"AC{row}"].Value = str.Meridians[2].G_as_Mother;
            sheet.Cells[$"AD{row}"].Value = str.Meridians[2].G_as_Master;
            sheet.Cells[$"AH{row}"].Value = str.Meridians[5].G_as_Mother;
            sheet.Cells[$"AI{row}"].Value = str.Meridians[5].G_as_Master;
            sheet.Cells[$"AM{row}"].Value = str.Meridians[8].G_as_Mother;
            sheet.Cells[$"AN{row}"].Value = str.Meridians[8].G_as_Master;
            sheet.Cells[$"AR{row}"].Value = str.Meridians[12].G_as_Mother;
            sheet.Cells[$"AS{row}"].Value = str.Meridians[12].G_as_Master;
            sheet.Cells[$"AW{row}"].Value = str.Meridians[4].G_as_Mother;
            sheet.Cells[$"AX{row}"].Value = str.Meridians[4].G_as_Master;
            sheet.Cells[$"BB{row}"].Value = str.Meridians[1].G_as_Mother;
            sheet.Cells[$"BC{row}"].Value = str.Meridians[1].G_as_Master;
            sheet.Cells[$"BG{row}"].Value = str.Meridians[7].G_as_Mother;
            sheet.Cells[$"BH{row}"].Value = str.Meridians[7].G_as_Master;
            sheet.Cells[$"BL{row}"].Value = str.Meridians[11].G_as_Mother;
            sheet.Cells[$"BM{row}"].Value = str.Meridians[11].G_as_Master;
            sheet.Cells[$"BQ{row}"].Value = str.Meridians[9].G_as_Mother;
            sheet.Cells[$"BR{row}"].Value = str.Meridians[9].G_as_Master;
            sheet.Cells[$"BV{row}"].Value = str.Meridians[10].G_as_Mother;
            sheet.Cells[$"BW{row}"].Value = str.Meridians[10].G_as_Master;


            // Set all Gray

            sheet.Cells[$"S{row}"]  .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"T{row}"]  .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"X{row}"]  .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"Y{row}"]  .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AC{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AD{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AH{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AI{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AM{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AN{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AR{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AS{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AW{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"AX{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BB{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BC{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BG{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BH{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BL{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BM{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BQ{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BR{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BV{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);
            sheet.Cells[$"BW{row}"] .Style.Font.Color.SetColor(1, 180, 180, 180);


            // Set for Corazon
            if (str.Meridians[6].G_as_Mother > maxLimit)
            {
                if (str.Meridians[6].I_Bioene < str.SonOf(str.Meridians[6]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"S{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"S{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[6].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[6].I_Bioene < str.SlaveOf(str.Meridians[6]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"T{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"T{row}"], CategoryColor.Dominancia);
            }
            // Set for Intestino Delgado
            if (str.Meridians[3].G_as_Mother > maxLimit)
            {

                if (str.Meridians[3].I_Bioene < str.SlaveOf(str.Meridians[3]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"X{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"X{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[3].G_as_Master > maxLimit)
            {

                if (str.Meridians[3].I_Bioene < str.SlaveOf(str.Meridians[3]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"Y{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"Y{row}"], CategoryColor.Dominancia);
            }
            // Set for Triple R
            if (str.Meridians[2].G_as_Mother > maxLimit)
            {

                if (str.Meridians[2].I_Bioene < str.SonOf(str.Meridians[2]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AC{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"AC{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[2].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[2].I_Bioene < str.SlaveOf(str.Meridians[2]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AD{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"AD{row}"], CategoryColor.Dominancia);
            }
            // Set for Pericardio
            if (str.Meridians[5].G_as_Mother > maxLimit)
            {

                if (str.Meridians[5].I_Bioene < str.SonOf(str.Meridians[5]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AH{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"AH{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[5].G_as_Master > maxLimit)
            {
               
                if (str.Meridians[5].I_Bioene < str.SlaveOf(str.Meridians[5]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AI{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"AI{row}"], CategoryColor.Dominancia);
            }
            // Set for Bazo
            if (str.Meridians[8].G_as_Mother > maxLimit)
            {
                
                if (str.Meridians[8].I_Bioene < str.SonOf(str.Meridians[8]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AM{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"AM{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[8].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[8].I_Bioene < str.SlaveOf(str.Meridians[8]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AN{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"AN{row}"], CategoryColor.Dominancia);
            }
            // Set for Estomago
            if (str.Meridians[12].G_as_Mother > maxLimit)
            {
                
                if (str.Meridians[12].I_Bioene < str.SonOf(str.Meridians[12]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AR{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"AR{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[12].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[12].I_Bioene < str.SlaveOf(str.Meridians[12]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AS{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"AS{row}"], CategoryColor.Dominancia);
            }
            // Set for Pulmon
            if (str.Meridians[4].G_as_Mother > maxLimit)
            {
                
                if (str.Meridians[4].I_Bioene < str.SonOf(str.Meridians[4]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AW{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"AW{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[4].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[4].I_Bioene < str.SlaveOf(str.Meridians[4]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"AX{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"AX{row}"], CategoryColor.Dominancia);
            }
            // Set for Intestion Grueso
            if (str.Meridians[1].G_as_Mother > maxLimit)
            {
                
                if (str.Meridians[1].I_Bioene < str.SonOf(str.Meridians[1]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BB{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"BB{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[1].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[1].I_Bioene < str.SlaveOf(str.Meridians[1]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BC{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"BC{row}"], CategoryColor.Dominancia);
            }
            // Set for Rinon
            if (str.Meridians[7].G_as_Mother > maxLimit)
            {
                
                if (str.Meridians[7].I_Bioene < str.SonOf(str.Meridians[7]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BG{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"BG{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[7].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[7].I_Bioene < str.SlaveOf(str.Meridians[7]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BH{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"BH{row}"], CategoryColor.Dominancia);
            }
            // Set for Vejiga
            if (str.Meridians[11].G_as_Mother > maxLimit)
            {
                
                if (str.Meridians[11].I_Bioene < str.SonOf(str.Meridians[11]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BL{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"BL{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[11].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[11].I_Bioene < str.SlaveOf(str.Meridians[11]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BM{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"BM{row}"], CategoryColor.Dominancia);
            }
            // Set for Higado
            if (str.Meridians[9].G_as_Mother > maxLimit)
            {                
                if (str.Meridians[9].I_Bioene < str.SonOf(str.Meridians[9]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BQ{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"BQ{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[9].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[9].I_Bioene < str.SlaveOf(str.Meridians[9]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BR{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"BR{row}"], CategoryColor.Dominancia);
            }
            // Set for Vesicula Biliar
            if (str.Meridians[10].G_as_Mother > maxLimit)
            {
                
                if (str.Meridians[10].I_Bioene < str.SonOf(str.Meridians[10]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BV{row}"], CategoryColor.Contracorriente);
                else
                    SetCategoryColor(sheet.Cells[$"BV{row}"], CategoryColor.Bloqueo);
            }
            if (str.Meridians[10].G_as_Master > maxLimit)
            {
                
                if (str.Meridians[10].I_Bioene < str.SlaveOf(str.Meridians[10]).I_Bioene)
                    SetCategoryColor(sheet.Cells[$"BW{row}"], CategoryColor.Contradominancia);
                else
                    SetCategoryColor(sheet.Cells[$"BW{row}"], CategoryColor.Dominancia);
            }
        }

        /// <summary>
        /// Ing G extraction
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="patientCode"></param>
        /// <param name="row"></param>
        public void Step4(ExcelWorksheet sheet, string patientCode, int row)
        {
            sheet.Cells[$"U{row}"].Value  = str.IndGOf(str.Meridians[6]);
            sheet.Cells[$"Z{row}"].Value  = str.IndGOf(str.Meridians[3]);
            sheet.Cells[$"AE{row}"].Value = str.IndGOf(str.Meridians[2]);
            sheet.Cells[$"AJ{row}"].Value = str.IndGOf(str.Meridians[5]);
            sheet.Cells[$"AO{row}"].Value = str.IndGOf(str.Meridians[8]);
            sheet.Cells[$"AT{row}"].Value = str.IndGOf(str.Meridians[12]);
            sheet.Cells[$"AY{row}"].Value = str.IndGOf(str.Meridians[4]);
            sheet.Cells[$"BD{row}"].Value = str.IndGOf(str.Meridians[1]);
            sheet.Cells[$"BI{row}"].Value = str.IndGOf(str.Meridians[7]);
            sheet.Cells[$"BN{row}"].Value = str.IndGOf(str.Meridians[11]);
            sheet.Cells[$"BS{row}"].Value = str.IndGOf(str.Meridians[9]);
            sheet.Cells[$"BX{row}"].Value = str.IndGOf(str.Meridians[10]);
        }

        public void Step5(ExcelWorksheet sheet, string patientCode, int row)
        {
            #region Load Dataset
            BimetOneReader bor = new BimetOneReader(Path.Combine(Dir, $"{patientCode}CANA.1"), BimetOneReader.FileFormat.NumeratedSessions);
            int lastSession = bor.Sessions[bor.Sessions.Length - 1];
            foreach (var org in str.Meridians)
            {
                decimal accumR = 0, accumL = 0;
                for (int session = 1; session <= lastSession; session++)
                {
                    accumL += Convert.ToDecimal(bor[session, org.Key - 1]);
                    accumR += Convert.ToDecimal(bor[session, org.Key - 1 + 12]);
                }

                org.Value.RightPotential = accumR / lastSession;
                org.Value.LeftPotential = accumL / lastSession;
            }
            #endregion

            #region Populate Excel

            var sheetEnergy = sheet.Workbook.Worksheets["Energy"];
            if (sheetEnergy == null) sheetEnergy = sheet.Workbook.Worksheets.Add("Energy");


            if (string.IsNullOrWhiteSpace(sheetEnergy.Cells[$"C{HeadRow}"].Value?.ToString()))
            {
                sheetEnergy.Cells[$"C{HeadRow}"].Value = "Codigo";
                FormatHead(sheetEnergy.Cells[$"C{HeadRow}"]);
            }
            if (string.IsNullOrWhiteSpace(sheetEnergy.Cells[$"D{HeadRow}"].Value?.ToString()))
            {
                sheetEnergy.Cells[$"D{HeadRow}"].Value = "Nombre";
                FormatHead(sheetEnergy.Cells[$"D{HeadRow}"]);
            }

            sheetEnergy.Cells[$"C{row}"].Value = patientCode;
            sheetEnergy.Cells[$"D{row}"].Value = sheet.Cells[$"D{row}"].Value;
            
            int excelColumn = 5; // Starting in Column 5 (A,B,C,D,E)
            foreach (var org in str.Meridians)
            {
                if (string.IsNullOrWhiteSpace(sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn)}{HeadRow}"].Value?.ToString()))
                {
                    sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn)}{HeadRow}"].Value = "I-" + org.Value.ShortName;
                    FormatHead(sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn)}{HeadRow}"]);
                }

                if (string.IsNullOrWhiteSpace(sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 13)}{HeadRow}"].Value?.ToString()))
                {
                    sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 13)}{HeadRow}"].Value = "R-" + org.Value.ShortName;
                    FormatHead(sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 13)}{HeadRow}"]);
                }

                if (string.IsNullOrWhiteSpace(sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 26)}{HeadRow}"].Value?.ToString()))
                {
                    sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 26)}{HeadRow}"].Value = org.Value.ShortName;
                    FormatHead(sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 26)}{HeadRow}"]);
                }

                sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn)}{row}"].Value = org.Value.LeftPotential;
                sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 13)}{row}"].Value = org.Value.RightPotential;
                sheetEnergy.Cells[$"{GetColNameFromIndex(excelColumn + 26)}{row}"].Formula = $"=({GetColNameFromIndex(excelColumn)}{row}+{GetColNameFromIndex(excelColumn + 13)}{row})/2";
                
                excelColumn++;
            }

            #endregion
        }

        private void FormatHead(ExcelRange excelRange)
        {
            excelRange.Style.Font.Color.SetColor(1, 0, 0, 0);
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            excelRange.Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);
            excelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;            
        }
    }
}
