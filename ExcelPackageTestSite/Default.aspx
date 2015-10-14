<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>
<%@ Register Assembly="ExportableExcelPackage" TagPrefix="xlp" Namespace="ExportableExcelPackage" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>ASP.NET</h1>
        <p class="lead">ASP.NET is a free web framework for building great Web sites and Web applications using HTML, CSS, and JavaScript.</p>
        <p><a href="http://www.asp.net" class="btn btn-primary btn-lg">Learn more &raquo;</a></p>
    </div>

    <div class="row">
        <div class="col-md-4">
            <xlp:ExcelPackage runat="server" ID="ExcelPackage" RenderHtml="true">
                <Worksheets>
                    <xlp:Worksheet runat="server" ID="Worksheet1" SheetName="Weekly">
                        <Items>
                            <xlp:WorksheetRow runat="server">
                                <Items>
                                    <xlp:NumberCell runat="server" Value="1" />
                                    <xlp:NumberCell runat="server" Value="2" /> 
                                </Items>
                            </xlp:WorksheetRow>
                            <xlp:NumberCell runat="server" Value="4" />
                            <xlp:NumberCell runat="server" Value="5" />
                            <xlp:NumberCell runat="server" Value="3" />
                        </Items>
                    </xlp:Worksheet>
                    <xlp:Worksheet runat="server" ID="Worksheet2" SheetName="Monthly">
                        <Items>
                            <xlp:WorksheetRow runat="server">
                                <Items>
                                    <xlp:NumberCell runat="server" Value="6" NumberFormat="#.00" />
                                    <xlp:NumberCell runat="server" Value="7" /> 
                                </Items>
                            </xlp:WorksheetRow>
                            <xlp:WorksheetRow runat="server">
                                <Items>
                                    <xlp:NumberCell runat="server" Value="8" />
                                    <xlp:NumberCell runat="server" Value="9" /> 
                                    <xlp:NumberCell runat="server" Value="10" /> 
                                </Items>
                            </xlp:WorksheetRow>
                        </Items>
                    </xlp:Worksheet>
                </Worksheets>
            </xlp:ExcelPackage>
            <asp:Button runat="server" ID="ExportBtn" OnClick="ExportBtn_Click" Text="Export to Excel" CssClass="btn btn-default"/>
            <h2>Getting started</h2>
            <p>
                ASP.NET Web Forms lets you build dynamic websites using a familiar drag-and-drop, event-driven model.
            A design surface and hundreds of controls and components let you rapidly build sophisticated, powerful UI-driven sites with data access.
            </p>
            <p>
                <a class="btn btn-default" href="http://go.microsoft.com/fwlink/?LinkId=301948">Learn more &raquo;</a>
            </p>
        </div>
        <div class="col-md-4">
            <h2>Get more libraries</h2>
            <p>
                NuGet is a free Visual Studio extension that makes it easy to add, remove, and update libraries and tools in Visual Studio projects.
            </p>
            <p>
                <a class="btn btn-default" href="http://go.microsoft.com/fwlink/?LinkId=301949">Learn more &raquo;</a>
            </p>
        </div>
        <div class="col-md-4">
            <h2>Web Hosting</h2>
            <p>
                You can easily find a web hosting company that offers the right mix of features and price for your applications.
            </p>
            <p>
                <a class="btn btn-default" href="http://go.microsoft.com/fwlink/?LinkId=301950">Learn more &raquo;</a>
            </p>
        </div>
    </div>
</asp:Content>
