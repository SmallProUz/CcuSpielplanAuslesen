using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.IO.Compression;

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

            if (basis.CreateCsv)
            {
                using (TextWriter writer = File.CreateText(basis.TargetPath + "Team" + basis.TeamNumber.ToString() + "Spieldaten.csv"))
                {
                    writer.WriteLine("subject;startdate;starttime;enddate;endtime"); //Titelzeile fuer importierbare Datei in Kalender
                    if (gefundeneDaten.Count > 0)
                    {
                        foreach (var item in gefundeneDaten)
                        {
                            writer.WriteLine("Curling " + item.Name + ";" + item.Day + "." + item.Month + "." + item.Year + ";" + item.StartHour + "." + item.StartMinute + ";"
                                + item.Day + "." + item.Month + "." + item.Year + ";" + item.EndHour + "." + item.StartMinute + ";");
                        }
                    }
                    else
                    {
                        writer.WriteLine("Fuer dieses Team wurden keine Daten gefunden.");
                    }
                }
            }
            

            if (basis.CreateIcs)
            {
                var path = basis.TargetPath + "\\Team" + basis.TeamNumber.ToString() + "Kalenderdateien";
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
                var dir = Directory.CreateDirectory(path);
                foreach (var curlingDatum in gefundeneDaten) //erzeuge pro Datum ein ICS file
                {
                    using (TextWriter writer = File.CreateText(path + "\\" + curlingDatum.Year + AddLeading0IfNeeded(curlingDatum.Month) + AddLeading0IfNeeded(curlingDatum.Day) + "-Match" + curlingDatum.Name + ".ics"))
                    {
                        writer.WriteLine("BEGIN:VCALENDAR");
                        writer.WriteLine("VERSION:2.0");
                        writer.WriteLine("PRODID:-//icalendar.org//CCUSpielplan//DE");
                        writer.WriteLine("BEGIN:VEVENT");
                        writer.WriteLine("CREATED:" + DateTime.Now.ToString("yyyyMMdd") + "T" + DateTime.Now.ToString("HHmmss") + "Z");
                        writer.WriteLine("DTSTAMP:" + DateTime.Now.ToString("yyyyMMdd") + "T" + DateTime.Now.ToString("HHmmss") + "Z");
                        writer.WriteLine("DTSTART:" + curlingDatum.Year + AddLeading0IfNeeded(curlingDatum.Month) + AddLeading0IfNeeded(curlingDatum.Day) + "T" + curlingDatum.StartHour + curlingDatum.StartMinute + "00");
                        writer.WriteLine("DTEND:" + curlingDatum.Year + AddLeading0IfNeeded(curlingDatum.Month) + AddLeading0IfNeeded(curlingDatum.Day) + "T" + curlingDatum.EndHour + curlingDatum.StartMinute + "00");
                        writer.WriteLine("LOCATION:" + "Curlinghalle Uzwil");
                        writer.WriteLine("UID:" + (curlingDatum.Year + AddLeading0IfNeeded(curlingDatum.Month) + AddLeading0IfNeeded(curlingDatum.Day) + "T" + curlingDatum.StartHour + curlingDatum.StartMinute + "00") + (curlingDatum.Year + AddLeading0IfNeeded(curlingDatum.Month) + AddLeading0IfNeeded(curlingDatum.Day) + "T" + curlingDatum.EndHour + curlingDatum.StartMinute + "00") + "|CCU");
                        writer.WriteLine("SUMMARY;LANGUAGE=en-us:" + "Curlingmatch " + curlingDatum.Name);
                        writer.WriteLine("DESCRIPTION:");
                        writer.WriteLine("X-ALT-DESC;FMTTYPE=text/html:");
                        writer.WriteLine("PRIORITY:5");
                        writer.WriteLine("SEQUENCE:0");
                        writer.WriteLine("STATUS:CONFIRMED");
                        writer.WriteLine("TRANSP:OPAQUE");
                        writer.WriteLine("CLASS:PUBLIC");
                        writer.WriteLine("X-MS-OLK-ALLOWEXTERNCHECK:FALSE");
                        writer.WriteLine("X-MS-OLK-AUTOFILLLOCATION:FALSE");
                        writer.WriteLine("X-MICROSOFT-CDO-ALLDAYEVENT:FALSE");
                        writer.WriteLine("X-MICROSOFT-DISALLOW-COUNTER:TRUE");
                        writer.WriteLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE");
                        writer.WriteLine("ORGANIZER;CN=\"Curling Club Uzwil\":MAILTO:m.louis.ch@gmail.com");
                        writer.WriteLine("BEGIN:VALARM");
                        writer.WriteLine("TRIGGER:-PT15M");
                        writer.WriteLine("ACTION:DISPLAY");
                        writer.WriteLine("DESCRIPTION:Reminder");
                        writer.WriteLine("END:VALARM");
                        writer.WriteLine("END:VEVENT");
                        writer.WriteLine("END:VCALENDAR");
                    }
                }
                ZipFile.CreateFromDirectory(path, path + ".zip");
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
        private static string AddLeading0IfNeeded(string input)
        {
            string output = "";
            if (input.Length < 2)
            {
                output = "0" + input;
            }
            else
            {
                output = input;
            }
            return output;
        }
    }
}
