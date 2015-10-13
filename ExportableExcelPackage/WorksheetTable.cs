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
    [DefaultProperty("Grid")]
    [ParseChildren(true)]
    [PersistChildren(true)]
    [ToolboxData("<{0}:Table runat=server></{0}:Table>")]
    public class WorksheetTable : WorksheetItem, IDataBoundControl
    {
        /// <summary>
        /// Need an item that is CompositeDataBoundControl, and don't really want to make WorksheetItem extend it.
        /// </summary>
        public WorksheetTableGrid Grid { get; set; }

        /// <summary>
        /// The collection of columns contained in this table.
        /// </summary>
        public WorksheetRow Columns {
            get
            {
                return Grid.Columns;
            }
            set
            {
                Grid.Columns = value;
            }
        }

        public virtual TableRowCollection Rows
        {
            get
            {
                return Grid.Rows;
            }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {

        }

        public override ExcelCellAddress AddItemToWorksheet(ExcelWorksheet worksheet)
        {
            foreach (WorksheetItem column in Columns.Items)
            {
                
            }
            throw new NotImplementedException();
        }

        public string[] DataKeyNames
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

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
            get { return Grid.DataSourceObject }
        }
    }
}
