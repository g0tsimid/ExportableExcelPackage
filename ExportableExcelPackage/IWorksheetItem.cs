using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExportableExcelPackage
{
    interface IWorksheetItem
    {
        /// <summary>
        /// The 0-based row number for the top-left position of the item, if applicable.
        /// </summary>
        int Row { get; set; }
        /// <summary>
        /// The 0-based column number for the top-left position of the item, if applicable.
        /// </summary>
        int Column { get; set; }

        /// <summary>
        /// Render this item to the Excel worksheet object. Return the end address after insertion.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns>The end ExcelCellAddress after adding this item to the worksheet.</returns>
        ExcelCellAddress AddItemToWorksheet(ExcelWorksheet worksheet);
    }
}
