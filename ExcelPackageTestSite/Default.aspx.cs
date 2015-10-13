using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void ExportBtn_Click(object sender, EventArgs e)
    {
        Response.ClearContent();

        Response.AddHeader("content-disposition", ExcelPackage.ContentDisposition);
        Response.ContentType = ExcelPackage.ContentType;
        Response.BinaryWrite(ExcelPackage.ExportAsByteArray());

        Response.End();
    }
}