using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Entities
{
    public class ExportExcel<TData> where TData : class, new()
    {
        public ExportExcel()
        {
            Detaills = new HashSet<TData>();
            WorkSheet = "Hoja";
        }

        public ExportExcel(IEnumerable<TData> detaills)
        {
            Detaills = detaills;
            WorkSheet = "Hoja";
        }

        public ExportExcel(string worksheetName, IEnumerable<TData> detaills)
        {
            WorkSheet = worksheetName;
            Detaills = detaills;
        }

        public string WorkSheet { get; set; }
        public object Header { get; set; }
        public object Footer { get; set; }
        public IEnumerable<TData> Detaills { get; set; }
    }
}
