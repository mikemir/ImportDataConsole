using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Exceptions
{
    public class NotFoundWorksheetExportException : Exception
    {
        public NotFoundWorksheetExportException(string name) :
            base($"La hoja \"{name}\" especificada no se encuentra en la plantilla de excel.")
        {
        }
    }
}
