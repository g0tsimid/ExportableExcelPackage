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
    [ToolboxData("<{0}:TextCell runat=server></{0}:TextCell>")]
    public class TextCell : WorksheetItem
    {
        [Bindable(true)]
        [Category("Appearance")]
        [DefaultValue(null)]
        [Localizable(true)]
        public decimal? Text
        {
            get
            {
                return ViewState["Text"] as decimal?;
            }

            set
            {
                ViewState["Text"] = value;
            }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {
            output.Write(Text);
        }


        public override ExcelCellAddress AddItemToWorksheet(ExcelWorksheet worksheet)
        {
            var cell = worksheet.Cells[Row + 1, Column + 1];
            cell.Value = Text;
            return cell.End;
        }
    }
}
