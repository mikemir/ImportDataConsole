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

        public ExportExcel(IEnumerable<TData> data)
        {
            Detaills = data;
            WorkSheet = "Hoja";
        }

        public string WorkSheet { get; set; }
        public object Header { get; set; }
        public object Footer { get; set; }
        public IEnumerable<TData> Detaills { get; set; }
    }
}
