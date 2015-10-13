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
    [ToolboxData("<{0}:WorksheetTableGrid runat=server></{0}:WorksheetTableGrid>")]
    public class WorksheetTableGrid : CompositeDataBoundControl, INamingContainer
    {
        protected Table table = new Table();
        /// <summary>
        /// The collection of columns contained in this grid.
        /// </summary>
        public WorksheetRow Columns { get; set; }

        public virtual TableRowCollection Rows
        {
            get
            {
                return table.Rows;
            }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {
            output.Write(Text);
        }
    }
}
