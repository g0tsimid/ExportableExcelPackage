using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExportableExcelPackage
{
    /// <summary>
    /// A base interface for an item than can appear on a worksheet.
    /// </summary>
    public abstract class WorksheetItem : WebControl
    {
        /// <summary>
        /// The 0-based row number for the top-left position of the item, if applicable.
        /// </summary>
        public virtual int Row { get; set; }
        /// <summary>
        /// The 0-based column number for the top-left position of the item, if applicable.
        /// </summary>
        public virtual int Column { get; set; }

        /// <summary>
        /// Render this item to the Excel worksheet object. Return the end address after insertion.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns>The end ExcelCellAddress after adding this item to the worksheet.</returns>
        public abstract ExcelCellAddress AddItemToWorksheet(ExcelWorksheet worksheet);
    }
}
