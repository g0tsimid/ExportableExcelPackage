using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using OfficeOpenXml;

namespace ExportableExcelPackage
{
    /// <summary>
    /// A row that defines a collection of other worksheet items on a given in the worksheet.
    /// </summary>
    [ParseChildren(true)]
    [PersistChildren(true)]
    public class WorksheetRow : WorksheetItem
    {
        private List<WorksheetItem> items;
        /// <summary>
        /// The row constitutes the entire x-axis of the spreadsheet, so its upper-left corner is (0, Row);
        /// </summary>
        public override int Column
        {
            get
            {
                return 0;
            }
            set
            {

            }
        }
        public override int Row { get; set; }
        [PersistenceMode(PersistenceMode.InnerProperty)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public List<WorksheetItem> Items
        {
            get
            {
                return items ?? (items = new List<WorksheetItem>());
            }
        }
        public override ExcelCellAddress AddItemToWorksheet(ExcelWorksheet worksheet)
        {
            int i = 0;
            ExcelCellAddress result = null;
            foreach (WorksheetItem item in Items)
            {
                item.Row = Row; // Inner cells should not have this defined.
                // Auto-generate column numbers, if not specified.
                item.Column = item.Column > 0 && item.Column != i ? item.Column : i;
                result = item.AddItemToWorksheet(worksheet);
                ++i;
            }
            return result;
        }


    }
}
