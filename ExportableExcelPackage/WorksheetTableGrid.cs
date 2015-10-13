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
        private List<WorksheetItem> columns;
        private List<WorksheetRow> rows;
        /// <summary>
        /// The collection of columns contained in this grid.
        /// </summary>
        public List<WorksheetItem> Columns {
            get
            {
                return columns ?? (columns = new List<WorksheetItem>());
            }
            set
            {
                columns = value;
            }
        }

        public virtual List<WorksheetRow> Rows
        {
            get
            {
                return rows ?? (rows = new List<WorksheetRow>());
            }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {
            
        }

        protected override int CreateChildControls(System.Collections.IEnumerable dataSource, bool dataBinding)
        {
            if (dataBinding)
            {
                Rows.Clear();
            }
            int count = 0;
            if (dataSource != null)
            {
                foreach (object dataItem in dataSource)
                {
                    WorksheetRow row = new WorksheetRow();
                    row.DataItem = dataItem;
                    foreach (WorksheetItem item in Columns)
                    {
                        item.DataItem = dataItem;
                        item.DataItemIndex = count++;
                        item.DisplayIndex = item.DataItemIndex;
                        row.Items.Add(item);
                        item.DataBind();
                    }
                }
            }
            
            return count;
        }
    }
}
