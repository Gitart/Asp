<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        body
        {
            font-family: Arial;
            font-size: 10pt;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    Search Customer:
    <asp:TextBox ID="txtSearch" runat="server"></asp:TextBox>
    <asp:Button ID="btnSearch" runat="server" Text="Search" />
    <hr />
    <asp:GridView ID="GridView1" HeaderStyle-BackColor="#3AC0F2" HeaderStyle-ForeColor="White"
        runat="server" AutoGenerateColumns="false" DataSourceID="GridDataSource" AllowPaging="true">
        <Columns>
            <asp:BoundField DataField="ContactName" HeaderText="Name" ItemStyle-Width="150" />
            <asp:BoundField DataField="City" HeaderText="City" ItemStyle-Width="150" />
            <asp:BoundField DataField="Country" HeaderText="Country" ItemStyle-Width="150" />
        </Columns>
    </asp:GridView>
    <asp:SqlDataSource ID="GridDataSource" runat="server" ConnectionString="<%$ ConnectionStrings:ConStr %>"
        SelectCommand="SELECT ContactName, City, Country FROM Customers" FilterExpression="ContactName LIKE '{0}%'">
        <FilterParameters>
            <asp:ControlParameter Name="Country" ControlID="txtSearch" PropertyName="Text" />
        </FilterParameters>
    </asp:SqlDataSource>
    </form>
</body>
</html>
