using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.ExcelHelper.Entities
{
    public class ImportExcel<TItem> where TItem : class, new()
    {
        public ImportExcel()
        {
            Item = new TItem();
        }

        public bool IsValid { get; set; }
        public string ValidationMessage { get; set; }
        public TItem Item { get; set; }
    }
}
