using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCUSpielplanAuslesen
{
    class StartFormContent
    {
        public string SourceFile { get; set; }
        public string TargetPath { get; set; }
        public int FirstRow { get; set; }
        public int FirstColumn { get; set; }
        public string StartDate { get; set; }
        public int TeamNumber { get; set; }
    }
}
