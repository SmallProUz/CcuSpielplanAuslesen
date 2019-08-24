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
                        if (line.StartsWith(basis.TeamNumber.ToString() + " -") || line.EndsWith("- " + basis.TeamNumber.ToString()))
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
                            neuesDatum.Name = line;
                            gefundeneDaten.Add(neuesDatum);
                        }
                    }
                }
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
    }
}
