using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OfficeOpenXml;

namespace ExportableExcelPackage
{
    /// <summary>
    /// A bindable table with named column definitions which can be configured to display
    /// conditionally.
    /// </summary>
    [ParseChildren(true)]
    [PersistChildren(true)]
    [DefaultProperty("Columns")]
    public class WorksheetTable : WorksheetItem, INamingContainer, IDataBoundControl
    {
        private WorksheetTableGrid grid;
        /// <summary>
        /// Need an item that is CompositeDataBoundControl, and don't really want to make WorksheetItem extend it.
        /// </summary>
        public WorksheetTableGrid Grid {
            get
            {
                return grid ?? (grid = new WorksheetTableGrid());
            }
            set
            {
                grid = value;
            }
        }

        /// <summary>
        /// The collection of columns contained in this table.
        /// </summary>
        [PersistenceMode(PersistenceMode.InnerDefaultProperty)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public List<WorksheetItem> Columns {
            get
            {
                return Grid.Columns;
            }
            set
            {
                Grid.Columns = value;
            }
        }

        /// <summary>
        /// The collection of rows contained in this table.
        /// </summary>
        public virtual List<WorksheetRow> Rows
        {
            get
            {
                return Grid.Rows;
            }
        }

        public override ExcelCellAddress AddItemToWorksheet(ExcelWorksheet worksheet)
        {
            DataBind();
            Grid.DataBind();
            int i = Row;
            ExcelCellAddress result = null;

            if (Rows.Count > 0)
            {
                // Add all rows from the grid to the worksheet.
                foreach (WorksheetItem item in Columns)
                {
                    // Render the header row first.
                }
                i++;
            }

            foreach (WorksheetRow row in Rows)
            {
                row.Row = i;
                row.Column = Column;
                result = row.AddItemToWorksheet(worksheet);
            }
            return result;
        }

        [Bindable(true)]
        [PersistenceMode(PersistenceMode.Attribute)]
        public string DataMember
        {
            get
            {
                return Grid.DataMember;
            }
            set
            {
                Grid.DataMember = value;
            }
        }

        object IDataBoundControl.DataSource
        {
            get
            {
                return Grid.DataSource;
            }
            set
            {
                Grid.DataSource = value;
            }
        }

        [Bindable(true)]
        [PersistenceMode(PersistenceMode.Attribute)]
        public string DataSourceID
        {
            get
            {
                return Grid.DataSourceID;
            }
            set
            {
                Grid.DataSourceID = value;
            }
        }

        public IDataSource DataSourceObject
        {
            get { return Grid.DataSourceObject; }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {

        }

        public string[] DataKeyNames { get; set; }
    }
}
