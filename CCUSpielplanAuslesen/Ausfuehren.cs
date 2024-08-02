using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCUSpielplanAuslesen
{
    class Ausfuehren
    {
        public static void Run(StartFormContent basis)
        {
            List<SpielDatumZeit> gefundeneDaten = new List<SpielDatumZeit>();
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(basis.SourceFile)))

            {

                var myWorksheet = xlPackage.Workbook.Worksheets[1];
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                int dateAddRow;
                int dateAddColumn;
                DateTime tempDateTime = new DateTime();

                for (int colNum = basis.FirstColumn; colNum <= totalColumns; colNum++)
                {
                    for (int rowNum = basis.FirstRow; rowNum <= totalRows; rowNum++)
                    {
                        var row = myWorksheet.Cells[rowNum, colNum, rowNum, colNum].Select(c => c.Value == null ? string.Empty : c.Value.ToString());//ganze linie
                        var line = string.Join(",", row);
                        if (EnthaeltZielnummerUndGegner(line,basis.TeamNumber)) //Hier ändern
                                               //if (line.StartsWith(basis.TeamNumber.ToString() + " -") || line.EndsWith("- " + basis.TeamNumber.ToString()))
                        {
                            SpielDatumZeit neuesDatum = new SpielDatumZeit();
                            tempDateTime = DateTime.Parse(basis.StartDate);
                            dateAddRow = (rowNum - basis.FirstRow) / 8;
                            dateAddColumn = (colNum - basis.FirstColumn) * 7;
                            tempDateTime = tempDateTime.AddDays(dateAddColumn + dateAddRow);
                            neuesDatum.Year = tempDateTime.Year.ToString();
                            neuesDatum.Month = tempDateTime.Month.ToString();
                            neuesDatum.Day = tempDateTime.Day.ToString();
                            if (((rowNum - basis.FirstRow - (dateAddRow*8))/4)>=1)
                            {
                                neuesDatum.StartHour = "20";
                                neuesDatum.StartMinute = "15";
                                neuesDatum.EndHour = "22";
                            }
                            else
                            {
                                neuesDatum.StartHour = "18";
                                neuesDatum.StartMinute = "00";
                                neuesDatum.EndHour = "20";
                            }
                            neuesDatum.Name = ReplaceHundreds(ReplaceHundreds(line)); //muss zwei mal durchlaufen da es nur den ersten hunderter ersetzt, kann aber zwei drin haben
                            gefundeneDaten.Add(neuesDatum);
                        }
                    }
                }
                if (basis.AddFinals)
                {
                    SpielDatumZeit FreitagFinal = new SpielDatumZeit();
                    SpielDatumZeit SamstagFinal = new SpielDatumZeit();
                    tempDateTime = DateTime.Parse(basis.FridayFinalDate);
                    FreitagFinal.Year = tempDateTime.Year.ToString();
                    FreitagFinal.Month = tempDateTime.Month.ToString();
                    FreitagFinal.Day = tempDateTime.Day.ToString();
                    FreitagFinal.StartHour = basis.FridayFinalStartTimeText;
                    FreitagFinal.StartMinute = "00";
                    FreitagFinal.EndHour = "22";
                    FreitagFinal.Name = "Clubmeisterschaft";
                    gefundeneDaten.Add(FreitagFinal);
                    tempDateTime = DateTime.Parse(basis.SaturdayFinalDate);
                    SamstagFinal.Year = tempDateTime.Year.ToString();
                    SamstagFinal.Month = tempDateTime.Month.ToString();
                    SamstagFinal.Day = tempDateTime.Day.ToString();
                    SamstagFinal.StartHour = basis.SaturdayFinalStartTimeText;
                    SamstagFinal.StartMinute = "00";
                    SamstagFinal.EndHour = "20";
                    SamstagFinal.Name = "Clubmeisterschaft";
                    gefundeneDaten.Add(SamstagFinal);
                }
                basis.TeamNumber = ReplaceHundreds(basis.TeamNumber);
                if (basis.AddToOriginal)
                {
                    var resultat = xlPackage.Workbook.Worksheets;
                    bool pageExists = false;
                    foreach (var item in resultat)
                    {
                        if (item.Name == "Team " + basis.TeamNumber)
                        {
                            pageExists = true;
                        }
                    }
                    if (pageExists)
                    {
                        ExcelWorksheet existingSheet = xlPackage.Workbook.Worksheets["Team " + basis.TeamNumber];
                        var usedRows = existingSheet.Dimension.End.Row;
                        for (int i = 1; i < usedRows; i++)
                        {
                            existingSheet.DeleteRow(i);
                        }
                        existingSheet.Cells[1, 1].Value = "Datum";
                        existingSheet.Cells[1, 2].Value = "Startzeit";
                        existingSheet.Cells[1, 3].Value = "Spiel";
                        if (gefundeneDaten.Count > 0)
                        {
                            int rowToLoad = 2;
                            foreach (var item in gefundeneDaten)
                            {
                                existingSheet.Cells[rowToLoad, 1].Value = item.Day + "." + item.Month + "." + item.Year;
                                existingSheet.Cells[rowToLoad, 2].Value = item.StartHour + "." + item.StartMinute;
                                existingSheet.Cells[rowToLoad, 3].Value = item.Name;
                                rowToLoad++;
                            }
                        }
                    }
                    else
                    {
                        ExcelWorksheet newSheet = xlPackage.Workbook.Worksheets.Add("Team " + basis.TeamNumber);
                        newSheet.Cells[1, 1].Value = "Datum";
                        newSheet.Cells[1, 2].Value = "Startzeit";
                        newSheet.Cells[1, 3].Value = "Spiel";
                        if (gefundeneDaten.Count > 0)
                        {
                            int rowToLoad = 2;
                            foreach (var item in gefundeneDaten)
                            {
                                newSheet.Cells[rowToLoad, 1].Value = item.Day + "." + item.Month + "." + item.Year;
                                newSheet.Cells[rowToLoad, 2].Value = item.StartHour + "." + item.StartMinute;
                                newSheet.Cells[rowToLoad, 3].Value = item.Name;
                                rowToLoad++;
                            }
                        }
                    }
                    xlPackage.Save();
                }

            }

            using (TextWriter writer = File.CreateText(basis.TargetPath+"Team"+basis.TeamNumber.ToString()+"Spieldaten.csv"))
            {
                if (gefundeneDaten.Count > 0)
                {
                    foreach (var item in gefundeneDaten)
                    {
                        writer.WriteLine(item.Day + "." + item.Month + "." + item.Year + ";" + item.StartHour + "." + item.StartMinute + ";" + item.EndHour + "." + item.StartMinute + ";" + item.Name);
                    }
                }
                else
                {
                    writer.WriteLine("Fuer dieses Team wurden keine Daten gefunden.");
                }
            }
            
        }
        private static bool EnthaeltZielnummerUndGegner(string inputString, int gesuchteZahl)
        {
            if (!string.IsNullOrWhiteSpace(inputString) && inputString.Length>1)
            {
                int indexErsteZiffer = -1;
                int indexLetzteZiffer = -1;
                int indexErstesTrennzeichen = -1;
                int indexLetztesTrennzeichen = -1;
                int anzahlTrennzeichen = 0;
                int ersteZahlInt = 0;
                int zweiteZahlInt = 0;
                string ersteZahl = "";
                string zweiteZahl = "";
                char tempChar;
                for (int i = 0; i < inputString.Length; i++)
                {
                    tempChar = inputString[i];
                    if (char.IsDigit(tempChar))
                    {
                        if (indexErsteZiffer == -1)
                        {
                            indexErsteZiffer = i;
                        }
                        else
                        {
                            indexLetzteZiffer = i;
                        }
                    }
                }
                if (indexErsteZiffer == -1 || indexLetzteZiffer == -1)//es wurde nur eine einzige Ziffer (oder gar keine) im String gefunden, es ist also keine gültige Angabe für ein Spiel mit Gegner
                {
                    return false;
                }
                if (indexLetzteZiffer > 0 && inputString.Length > indexLetzteZiffer +1) //es hat noch weitere Zeichen im String nach der letzten Ziffer
                {
                    inputString = inputString.Substring(0, indexLetzteZiffer + 1); //Abschneiden der letzten Zeichen
                }
                if (indexErsteZiffer > 0) //es hat noch Zeichen vor der ersten Ziffer
                {
                    inputString = inputString.Substring(indexErsteZiffer); //Abschneiden der Zeichen an der Front
                }
                for (int i = 0; i < inputString.Length; i++)
                {
                    tempChar = inputString[i];
                    if (!char.IsDigit(tempChar))
                    {
                        if (indexErstesTrennzeichen == -1)
                        {
                            indexErstesTrennzeichen = i;
                        }
                        else
                        {
                            indexLetztesTrennzeichen = i;
                        }
                    }
                }
                if (indexErstesTrennzeichen < 0)
                {
                    return false; //es wurde kein Trennzeichen gefunden, also ist nur eine Zahl im String, also keine gültige Angabe
                }
                else
                {
                    if (indexLetztesTrennzeichen < 0)
                    {
                        anzahlTrennzeichen = 1;
                    }
                    else
                    {
                        anzahlTrennzeichen = indexLetztesTrennzeichen - indexErstesTrennzeichen + 1;
                    }
                }
                if (anzahlTrennzeichen > 0)
                {
                    ersteZahl = inputString.Substring(0, indexErstesTrennzeichen);
                    zweiteZahl = inputString.Substring(indexErstesTrennzeichen + anzahlTrennzeichen);
                    if (!int.TryParse(ersteZahl,out ersteZahlInt) || !int.TryParse(zweiteZahl,out  zweiteZahlInt))
                    {
                        return false;
                    }
                }
                if (ersteZahlInt == gesuchteZahl || zweiteZahlInt == gesuchteZahl)
                {
                    return true;
                }
            }
            return false;
        }
        private static string ReplaceHundreds(string inputString)
        {
            if (!string.IsNullOrEmpty(inputString)) 
            {
                if (inputString.Contains("101"))
                {
                    return inputString.Replace("101", "1");
                }

                if (inputString.Contains("102"))
                {
                    return inputString.Replace("102", "2");
                }

                if (inputString.Contains("103"))
                {
                    return inputString.Replace("103", "3");
                }

                if (inputString.Contains("104"))
                {
                    return inputString.Replace("104", "4");
                }

                if (inputString.Contains("105"))
                {
                    return inputString.Replace("105", "5");
                }

                if (inputString.Contains("106"))
                {
                    return inputString.Replace("106", "6");
                }

                if (inputString.Contains("107"))
                {
                    return inputString.Replace("107", "7");
                }

                if (inputString.Contains("108"))
                {
                    return inputString.Replace("108", "8");
                }

                if (inputString.Contains("109"))
                {
                    return inputString.Replace("109", "9");
                }
            }
            return inputString;
        }
        private static int ReplaceHundreds(int inputInt)
        {
            switch (inputInt)
            {
                case 101:
                    return 1;
                case 102:
                    return 2;
                case 103:
                    return 3;
                case 104:
                    return 4;
                case 105:
                    return 5;
                case 106:
                    return 6;
                case 107:
                    return 7;
                case 108:
                    return 8;
                case 109:
                    return 9;
                default: 
                    return inputInt;
            }
        }
    }
}
