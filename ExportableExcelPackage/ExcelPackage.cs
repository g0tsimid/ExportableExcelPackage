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
    /// <summary>
    /// A wrapper for the EPPlus ExcelPackage that allows the constructed package
    /// to be configured through aspx markup and data sources.
    /// </summary>
    [DefaultProperty("RenderHtml")]
    [ParseChildren(true)]
    [PersistChildren(true)]
    [ToolboxData("<{0}:ExcelPackage runat=server></{0}:ExcelPackage>")]
    public class ExcelPackage : WebControl, INamingContainer
    {
        const string DEFAULT_FILENAME = "Workbook.xlsx";
        const string DEFAULT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        private WorksheetCollection worksheets;

        /// <summary>
        /// Indicate whether the component should render any html to the page. Its default use case
        /// is to be embedded on the page and rendered using ExportAsByteArray();
        /// </summary>
        [Bindable(true)]
        [Category("Appearance")]
        [DefaultValue(false)]
        [Description("Indicate whether the component should render any HTML to the page.")]
        public bool RenderHtml
        {
            get
            {
                return (bool?)ViewState["RenderHtml"] ?? false;
            }
            set
            {
                ViewState["RenderHtml"] = value;
            }
        }

        /// <summary>
        /// The file name for the exported excel package.
        /// </summary>
        [Bindable(true)]
        [Category("HTTP")]
        [DefaultValue(DEFAULT_FILENAME)]
        [Description("The file name for the exported excel package.")]
        public string FileName
        {
            get
            {
                return (string)ViewState["FileName"] ?? DEFAULT_FILENAME;
            }
            set
            {
                ViewState["FileName"] = value;
            }
        }

        /// <summary>
        /// The HTTP content type for the byte string returned by ExportAsByteString.
        /// </summary>
        [Bindable(true)]
        [Category("HTTP")]
        [DefaultValue(DEFAULT_CONTENT_TYPE)]
        [Description("The response content type for the excel document.")]
        public string ContentType
        {
            get
            {
                return (string)ViewState["ContentType"] ?? DEFAULT_CONTENT_TYPE;
            }
            set{
                ViewState["ContentType"] = value;
            }
        }

        [PersistenceMode(PersistenceMode.InnerProperty)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public WorksheetCollection Worksheets
        {
            get
            {
                return worksheets ?? (worksheets = new WorksheetCollection());
            }
        }

        /// <summary>
        /// The HTTP content disposition for the package.
        /// </summary>
        public string ContentDisposition
        {
            get
            {
                return String.Format("attachment;filename=\"{0}\"", FileName);
            }
        }
 
        /// <summary>
        /// Render a representation of this ExcelPackage as HTML.
        /// </summary>
        /// <param name="output"></param>
        protected override void RenderContents(HtmlTextWriter output)
        {
            if (RenderHtml)
            {
                foreach (Worksheet worksheet in Worksheets)
                {
                    worksheet.RenderControl(output);
                }
            }
        }

        /// <summary>
        /// Export the contents of this ExcelPackage as a byte array.
        /// </summary>
        /// <returns></returns>
        public byte[] ExportAsByteArray()
        {
            OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage();

            foreach (Worksheet worksheet in worksheets)
            {
                OfficeOpenXml.ExcelWorksheet excelSheet;
                excelSheet = package.Workbook.Worksheets.Add(worksheet.SheetName);
                worksheet.AddItemsToWorksheet(excelSheet);
            }
            return package.GetAsByteArray();
        }
    }
}
