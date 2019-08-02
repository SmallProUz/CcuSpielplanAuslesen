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

                var myWorksheet = xlPackage.Workbook.Worksheets["Tabelle1"];
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
                           // if (line.IndexOf(basis.TeamNumber.ToString() + " -")>=0 || line.IndexOf("- " + basis.TeamNumber.ToString()) >= 0)
                        { //diese Entscheidung funktioniert nur mit Zahlen grösser gleich 10!
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

            }

            using (TextWriter writer = File.CreateText(basis.TargetPath+"Team"+basis.TeamNumber.ToString()+"Spieldaten.csv"))
            {
                foreach (var item in gefundeneDaten)
                {
                    writer.WriteLine(item.Day+"."+item.Month + "."+item.Year + ";"+item.StartHour + "."+item.StartMinute + ";"+item.EndHour + "."+item.StartMinute + ";"+item.Name);
                }
            }
            
        }
    }
}
