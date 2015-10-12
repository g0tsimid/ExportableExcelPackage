using System;
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
    [DefaultProperty("Text")]
    [ToolboxData("<{0}:NumberCell runat=server></{0}:NumberCell>")]
    public class NumberCell : WorksheetItem
    {
        [Bindable(true)]
        [Category("Appearance")]
        [DefaultValue("#")]
        public string NumberFormat
        {
            get
            {
                return (string) ViewState["NumberFormat"] ?? "#";
            }
            set
            {
                ViewState["NumberFormat"] = value;
            }
        }

        [Bindable(true)]
        [Category("Appearance")]
        [DefaultValue(null)]
        [Localizable(true)]
        public decimal? Value
        {
            get
            {
                return ViewState["Value"] as decimal?;
            }

            set
            {
                ViewState["Value"] = value;
            }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {
            output.Write(Value);
        }


        public override ExcelCellAddress AddItemToWorksheet(ExcelWorksheet worksheet)
        {
            var cell = worksheet.Cells[Row + 1, Column + 1];
            cell.Style.Numberformat.Format = NumberFormat;
            cell.Value = Value;
            return cell.End;
        }
    }
}
