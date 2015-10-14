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
        private List<WorksheetTableDataHeaderItem> columns;
        private List<WorksheetTableRow> rows;
        /// <summary>
        /// The collection of columns contained in this grid.
        /// </summary>
        public List<WorksheetTableDataHeaderItem> Columns
        {
            get
            {
                return columns ?? (columns = new List<WorksheetTableDataHeaderItem>());
            }
            set
            {
                columns = value;
            }
        }

        public virtual List<WorksheetTableRow> Rows
        {
            get
            {
                return rows ?? (rows = new List<WorksheetTableRow>());
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
                    WorksheetTableRow row = new WorksheetTableRow();
                    foreach (WorksheetTableDataHeaderItem item in Columns)
                    {
                        WorksheetTableDataItem rowItem = new WorksheetTableDataItem();

                        rowItem.DataItem = dataItem;
                        rowItem.DataItemIndex = count++;
                        rowItem.DisplayIndex = item.DataItemIndex;
                        rowItem.DataBind();
                        row.Add(rowItem);
                    }
                    Rows.Add(row);
                }
            }
            
            return count;
        }
    }
}
