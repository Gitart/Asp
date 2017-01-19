<%@ Page language="VB"
         Trace="false"
         title="Kaртка інінціатора"
         MaintainScrollPositionOnPostback=true
         inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,
         Microsoft.SharePoint,
         Version=12.0.0.0,
         Culture=neutral,
         PublicKeyToken=71e9bce111e9429c"
         Debug="true" %>

<%@ Import Namespace = "System.Data.Odbc" %>
<%@ Import Namespace = "System"  %>
<%@ Import Namespace = "System.IO"  %>
<%@ Import Namespace = "System.Data"  %>
<%@ Import Namespace = "System.IO.DirectoryInfo"  %>
<%@ Import Namespace = "System.IO.Directory"  %>
<%@ Import Namespace = "System.IO.File"  %>
<%@ Import Namespace = "System.Web.HttpServerUtility"  %>
<%@ Import Namespace = "System.Data"  %>
<%@ Import Namespace = "System.Web"  %>
<%@ Import Namespace = "System.Web.Sessionstate"  %>
<%@ Import Namespace = "System.Web.Configuration"  %>
<%@ Import Namespace = "System.Web.HttpApplication"  %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import Namespace = "System.Collections.Specialized" %>
<%@ Import Namespace = "System.Configuration" %>
<%@ Import Namespace = "System.Collections.ObjectModel" %>
<%@ Import Namespace = "System.Collections" %>
<%@ Import Namespace = "System.Text" %>
<%@ Import Namespace = "System.Web.UI.Page" %>
<%@ Import Namespace = "System.Web.HttpRequest" %>
<%@ Import Namespace = "System.Threading" %>
<%@ Import Namespace = "System.Globalization" %>
<%@ Import Namespace = "Microsoft.SharePoint.WebControls"  %>
<%@ Import Namespace = "System.Data.SqlClient" %>

<script id="SVB1" runat="server" type="text/vb" >

Public Iselect as Integer
Public Isdelete as Integer

'***************************************************************************************************************************************************
'Загрузка страницы
'***************************************************************************************************************************************************
Sub Page_load(ByVal sender As Object, ByVal e As EventArgs)

  Label1.text="Єкспорт в файл"
End sub

'Imports System.Data

  Protected Sub B111_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    Dim tw As New StringWriter()
    Dim hw As New System.Web.UI.HtmlTextWriter(tw)
    Dim frm As HtmlForm = new HtmlForm()

    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader("content-disposition", "attachment;filename=tabbs.html")
    Response.Charset = "windows-1251"
    EnableViewState = False
    Controls.Add(frm)
    frm.Controls.Add(D111)
    frm.RenderControl(hw)
    Response.Write(tw.ToString())
    Response.End()

End Sub


Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    GV.AllowPaging="False"
    GV.DataBind
    Dim tw As New StringWriter()
    Dim hw As New System.Web.UI.HtmlTextWriter(tw)
    Dim frm As HtmlForm = new HtmlForm()

    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader("content-disposition", "attachment;filename=sss.xls")
    Response.Charset = "windows-1251"
    EnableViewState = False
    Controls.Add(frm)
    frm.Controls.Add(GV)
    frm.RenderControl(hw)
    'Response.Write(tw.ToString())

    Response.Write("<h1>,kbyyyyyyyyyy</h1>")
    Response.End()

    GV.AllowPaging="True"
    GV.Databind
End Sub


    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
           '''

      End Sub

Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs)

   GV.AllowPaging="False"

   GV.DataBind

    Dim sw As New StringWriter()
    Dim hw As New System.Web.UI.HtmlTextWriter(sw)
    Dim frm As HtmlForm = New HtmlForm()

    Page.Response.AddHeader("content-disposition", "attachment;filename=Team.xls")
    Page.Response.ContentType = "application/vnd.ms-excel"
    Page.Response.Charset = "windows-1251"
    Page.Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1251")

    Page.EnableViewState = False
    frm.Attributes("runat") = "server"
    Controls.Add(frm)
    frm.Controls.Add(GV)
    frm.RenderControl(hw)
    Response.Write(sw.ToString())
    Response.End()

    GV.AllowPaging="True"
    GV.Databind

End Sub

Protected Sub BE_Click(ByVal sender As Object, ByVal e As System.EventArgs)

GV.AllowPaging="False"
GV.DataBind

    Dim sw As New StringWriter()
    Dim hw As New System.Web.UI.HtmlTextWriter(sw)
    Dim frm As HtmlForm = New HtmlForm()

   'context.Response.Charset = "windows-1251"
   'context.Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1251")
   'context.Response.ContentType = "text/html"
   'context.Response.Write("document.write(""ГУТ!"");"​ )

    Page.Response.AddHeader("content-disposition", "attachment;filename=FileName.doc")
    Page.Response.ContentType = "application/vnd.word"
    Page.Response.Charset = "windows-1251"
    Page.Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1251")

    Page.EnableViewState = False
    frm.Attributes("runat") = "server"
    Controls.Add(frm)
    frm.Controls.Add(GV)
    frm.RenderControl(hw)
    Response.Write(sw.ToString())
    Response.End()

    GV.AllowPaging="True"
    GV.Databind

End Sub

</script>

<html dir="ltr">

<head>
<META name="WebPartPageExpansion" content="full">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Untitled 16</title>
<meta name="Microsoft Theme" content="Winner 1011, default">
<style type="text/css">
.ms-simple3-main {
                    border: 1.5pt solid black;
}
.ms-simple3-tl {
                    font-weight: bold;
                    color: white;
                    border-style: none;
                    background-color: black;
}
.ms-simple3-left {
                    border-style: none;
}
.ms-simple3-top {
                    font-weight: bold;
                    color: white;
                    border-style: none;
                    background-color: black;
}
.ms-simple3-even {
                    border-style: none;
}
.style1 {
                    border: 2px solid #d0dfd6;
}
</style>
</head>

<body>

<form id="form1" runat="server">
<table style="width: 100%" class="style1">
                    <tr>
                                        <td style="width: 15px">&nbsp;</td>
                                        <td>
<asp:Button runat="server" Text="Export" id="btnExportExcel"
OnClick="Button1_Click" Width="95px"/>&nbsp;
<asp:Button runat="server" Text="Export - XLS" id="btnExcel2"
OnClick="btnExcel_Click" Width="95px"/>&nbsp;
<asp:Button runat="server" Text="Export - Word " id="bE2" OnClick="BE_Click"
Width="95px"/>&nbsp;

<asp:Button runat="server" Text="Export - TABLE " id="bE3" OnClick="B111_Click"
Width="106px"/>&nbsp;

<asp:Label runat="server" Text="Label" id="Label1"></asp:Label>
                                        </td>
                    </tr>
                    <tr>
                                        <td style="width: 15px">&nbsp;</td>
                                        <td>

<asp:GridView runat="server" id="GV"
              ForeColor="#333333"
              CellPadding="4"
              AutoGenerateColumns="False"
              DataSourceID="SqlDataSource1" GridLines="Both"
              Width="1242px"
              Height="181px"

              AllowPaging="True">

                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <Columns>
                                        <asp:boundfield DataField="Fio" HeaderText="Fio" SortExpression="Fio">
                                        </asp:boundfield>
                                        <asp:boundfield DataField="Department" HeaderText="Отдел" SortExpression="Department">
                                        </asp:boundfield>
                                        <asp:boundfield DataField="Note" HeaderText="Note" SortExpression="Note">
                                        </asp:boundfield>
                                        <asp:boundfield DataField="Mai_Namel" HeaderText="Mai_Namel" ReadOnly="True" SortExpression="Mai_Namel">
                                        </asp:boundfield>
                                        <asp:boundfield DataField="Tel_mob" HeaderText="Tel_mob" SortExpression="Tel_mob">
                                        </asp:boundfield>
                                        <asp:boundfield DataField="Tel_dom" HeaderText="Tel_dom" SortExpression="Tel_dom">
                                        </asp:boundfield>
                                        <asp:boundfield DataField="Tel_vnutr" HeaderText="Tel_vnutr" SortExpression="Tel_vnutr">
                                        </asp:boundfield>
                    </Columns>
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
</asp:GridView>
<asp:SqlDataSource runat="server"
                   ID="SqlDataSource1"
                   ProviderName="System.Data.SqlClient"
                   ConnectionString="Data Source=g4;Initial Catalog=Winner;User ID=userrrrr;Password=0000000"

SelectCommand="SELECT [Fio], [Department], [Note], [Mai Namel] AS Mai_Namel, [Tel_mob], [Tel_dom], [Tel_vnutr], DEP FROM [A_Users] WHERE DEP=3">
</asp:SqlDataSource>

                                        </td>
                    </tr>
</table>
<br>

<table style="height: 39px; width: 86px" >

</table>

<br>
<p>
&nbsp;</p>
</form>

<table runat="server"  id="D111" style="width: 100%" class="ms-simple3-main">
                    <!-- fpstyle: 3,011111100 -->
                    <tr>
                                        <td style="width: 697px" class="ms-simple3-tl">
                                        SSSSS</td>
                                        <td class="ms-simple3-top">SSSSS</td>
                    </tr>
                    <tr>
                                        <td style="width: 697px" class="ms-simple3-left">
                                        SSSS</td>
                                        <td class="ms-simple3-even">SSS</td>
                    </tr>
                    <tr>
                                        <td style="width: 697px" class="ms-simple3-left">
                                        SSS</td>
                                        <td class="ms-simple3-even">DDFF</td>
                    </tr>
</table>

</body>

</html>
