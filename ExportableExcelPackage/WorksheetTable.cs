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
    [DefaultProperty("DataSource")]
    [ParseChildren(true)]
    [PersistChildren(true)]
    [ToolboxData("<{0}:WorksheetTable runat=server></{0}:WorksheetTable>")]
    public class WorksheetTable : WorksheetItem
    {
        [PersistenceMode(PersistenceMode.InnerProperty)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public WorksheetTableGrid Grid
        {
            get;
            set;
        }

        [PersistenceMode(PersistenceMode.InnerProperty)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public DataSourceControl DataSource { get; set; }

        protected override void RenderContents(HtmlTextWriter output)
        {
            Grid.RenderControl(output);
        }

        public override OfficeOpenXml.ExcelCellAddress AddItemToWorksheet(OfficeOpenXml.ExcelWorksheet worksheet)
        {
            int i = Row;
            Grid.DataSource = DataSource;
            Grid.DataBind();
            foreach (WorksheetTableRow row in Grid.Rows)
            {
                foreach (WorksheetTableDataItem item in row)
                {
                    item.AddItemToWorksheet(worksheet);
                }
            }
            return null;
        }
    }
}
