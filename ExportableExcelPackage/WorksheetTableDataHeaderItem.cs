using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportableExcelPackage
{
    public class WorksheetTableDataHeaderItem : WorksheetTableDataItem
    {
        [Bindable(true)]
        public string HeaderText { get; set; }
    }
}
