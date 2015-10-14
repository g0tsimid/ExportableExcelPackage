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
    [ToolboxData("<{0}:WorksheetTableRow runat=server></{0}:WorksheetTableRow>")]
    public class WorksheetTableRow : List<WorksheetTableDataItem>
    {
        
    }
}
