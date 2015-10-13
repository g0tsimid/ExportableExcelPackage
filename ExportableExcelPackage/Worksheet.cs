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
    /// <summary>
    /// A single worksheet that comprises part of an excel workbook.
    /// </summary>
    [ParseChildren(true)]
    [PersistChildren(true)]
    public class Worksheet : WebControl, INamingContainer
    {
        private WorksheetItemCollection items;
        /// <summary>
        /// The name that should be assigned to this worksheet.
        /// </summary>
        public string SheetName
        {
            get
            {
                return (string)ViewState["SheetName"];
            }
            set
            {
                ViewState["SheetName"] = value;
            }
        }
        /// <summary>
        /// The collection of items on this worksheet.
        /// </summary>
        [PersistenceMode(PersistenceMode.InnerProperty)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public WorksheetItemCollection Items {
            get
            {
                return items ?? (items = new WorksheetItemCollection());
            }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {
            throw new NotImplementedException();
        }

        public ExcelCellAddress AddItemsToWorksheet(ExcelWorksheet worksheet)
        {
            ExcelCellAddress result = null;
            int i = 0;
            // Add each individual item to the worksheet.
            foreach (WorksheetItem item in Items)
            {
                item.Row = item.Row > 0 && i != item.Row ? item.Row : i;
                result = item.AddItemToWorksheet(worksheet);
                i++;
            }
            // The final address
            return result;
        }
    }
}
