using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.ExcelHelper.Entities
{
    public class ExportExcel<TData> where TData : class, new()
    {
        public ExportExcel()
        {
            Data = new HashSet<TData>();
        }

        public string WorkSheet { get; set; }
        public object Header { get; set; }
        public object Footer { get; set; }
        public IEnumerable<TData> Data { get; set; }
    }
}
