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
    [ParseChildren(true)]
    [PersistChildren(true)]
    public class Worksheet : WebControl
    {
        private ItemCollection items;
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

        [PersistenceMode(PersistenceMode.InnerProperty)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public ItemCollection Items {
            get
            {
                return items ?? (items = new ItemCollection());
            }
        }

        protected override void RenderContents(HtmlTextWriter output)
        {
            
        }

        public void AddItemsToWorksheet(ExcelWorksheet worksheet)
        {
            foreach (WorksheetItem item in Items)
            {
                item.AddItemToWorksheet(worksheet);
            }
        }
    }
}
