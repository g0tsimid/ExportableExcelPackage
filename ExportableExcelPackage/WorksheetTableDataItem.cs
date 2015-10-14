using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExportableExcelPackage
{
    [DefaultProperty("Text")]
    [ToolboxData("<{0}:WorksheetTableDataItem runat=server></{0}:WorksheetTableDataItem>")]
    public class WorksheetTableDataItem : WebControl, IDataItemContainer, IWorksheetItem
    {
        public object DataItem { get; set; }
        public int DataItemIndex { get; set; }
        public int DisplayIndex { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }

        public OfficeOpenXml.ExcelCellAddress AddItemToWorksheet(OfficeOpenXml.ExcelWorksheet worksheet)
        {
            throw new NotImplementedException();
        }
    }
}
