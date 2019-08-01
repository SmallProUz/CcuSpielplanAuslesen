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


                for (int colNum = basis.FirstColumn; colNum <= totalColumns; colNum++)
                {
                    for (int rowNum = basis.FirstRow; rowNum <= totalRows; rowNum++)
                    {
                        var row = myWorksheet.Cells[rowNum, colNum, rowNum, colNum].Select(c => c.Value == null ? string.Empty : c.Value.ToString());//ganze linie
                        var line = string.Join(",", row);
                        if (line.IndexOf(basis.TeamNumber.ToString() + " -")>=0 || line.IndexOf("- " + basis.TeamNumber.ToString()) >= 0)
                        { //diese Entscheidung funktioniert nur mit Zahlen grösser gleich 10!
                            SpielDatumZeit neuesDatum = new SpielDatumZeit();
                            gefundeneDaten.Add(neuesDatum);
                        }
                    }
                }

            }
        }
    }
}
