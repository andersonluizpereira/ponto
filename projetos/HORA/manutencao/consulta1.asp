<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:GridView ID="GridView1" runat="server" AllowSorting="True" 
            AutoGenerateColumns="False" BackColor="White" BorderColor="#CCCCCC" 
            BorderStyle="None" BorderWidth="50px" CellPadding="3" DataKeyNames="User_ID" 
            DataSourceID="SqlDataSource1" PageSize="20" Width="857px">
            <FooterStyle BackColor="White" ForeColor="#000066" />
            <RowStyle ForeColor="#000066" />
            <Columns>
                <asp:BoundField DataField="DT_DATA" HeaderText="DT_DATA" 
                    SortExpression="DT_DATA" />
                <asp:BoundField DataField="User_ID" HeaderText="User_ID" ReadOnly="True" 
                    SortExpression="User_ID" />
                <asp:BoundField DataField="Nome" HeaderText="Nome" SortExpression="Nome" />
                <asp:BoundField DataField="HR_ENTRADA" HeaderText="HR_ENTRADA" 
                    SortExpression="HR_ENTRADA" />
                <asp:BoundField DataField="HR_SAIDA" HeaderText="HR_SAIDA" 
                    SortExpression="HR_SAIDA" />
            </Columns>
            <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
            <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BorderStyle="Dotted" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:ConnectionString1 %>" 
            ProviderName="<%$ ConnectionStrings:ConnectionString1.ProviderName %>" 
            SelectCommand="SELECT DISTINCT [DT_DATA], [User_ID], [Nome], [HR_ENTRADA], [HR_SAIDA] FROM [ConsultaHoraDiaria] ORDER BY [Nome]">
        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
