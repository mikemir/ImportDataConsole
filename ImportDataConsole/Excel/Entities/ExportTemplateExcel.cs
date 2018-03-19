using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Entities
{
    public class TemplateData
    {
        private readonly bool _isTable;

        public TemplateData(string searchKey, string value)
        {
            SearchKey = searchKey;
            Value = value;
        }

        public TemplateData(string searchKey, IEnumerable<object> value)
        {
            SearchKey = searchKey;
            Value = value;
            _isTable = true;
        }

        public string SearchKey { get; set; }
        public object Value { get; set; }

        public bool IsTable { get { return _isTable; } }

        public IEnumerable<object> GetDataTable()
        {
            return _isTable ? (IEnumerable<object>)Value : null;
        }
    }

    public class ExportTemplateExcel
    {
        public ExportTemplateExcel(string worksheetName, IEnumerable<TemplateData> detaills)
        {
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentNullException(nameof(worksheetName));

            if(detaills == null)
                throw new ArgumentNullException(nameof(detaills));

            WorkSheet = worksheetName;
            Detaills = detaills;
        }

        public string WorkSheet { get; set; }
        public IEnumerable<TemplateData> Detaills { get; set; }
    }
}
