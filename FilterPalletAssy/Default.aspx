<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="FilterPalletAssy.WebForm1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Filter Pallet Assembly</title>
     <link href="~/Styles/Main.css" rel="stylesheet" type="text/css" /> 
</head>
<body>
    <form id="form1" DefaultButton="btnCreate" runat="server">
    <div>
        <asp:Label ID="lblTItle" CssClass="lblTitle" runat="server" Text="FILTER PALLET MANAGEMENT"></asp:Label>
        <asp:Label ID="lblError" CssClass="lblError" runat="server" Text=""></asp:Label>
        <asp:Label ID="lblUserName" CssClass="lblUserName" runat="server" Text="" Visible="false" ></asp:Label>
        <asp:Button ID="btnCreate" CssClass="btnCreate" runat="server" Text="Create" />
        <asp:Button ID="btnResume" CssClass="btnResume" runat="server" Text="Resume" />
        <asp:Button ID="btnReset" CssClass="btnReset" runat="server" Text="Reset / Close" CausesValidation="false" />
        <asp:DropDownList ID="ddlMode" CssClass="ddlMode" runat="server" Visible="false" AutoPostBack="true">
            <asp:ListItem Value="0">-SELECT MODE-</asp:ListItem>
            <asp:ListItem Value="1">Assembly</asp:ListItem>
            <asp:ListItem Value="2">Spares</asp:ListItem>
            <asp:ListItem Value="3">QAR</asp:ListItem>
            <asp:ListItem Value="4">LAB</asp:ListItem>
            <asp:ListItem Value="5">In Process QA</asp:ListItem>
        </asp:DropDownList>
       <%-- <asp:DropDownList ID="ddlPrinters" CssClass="ddlPrinters" runat="server" AutoPostBack="true" Visible="false" />
        <asp:Label ID="lblDefaultPersonalPrinter" CssClass="lblDefaultPersonalPrinter" runat="server" Text="Make Printer Personal Default?" Visible="false"></asp:Label>
        <asp:CheckBox ID="chkDefaultPersonal" CssClass="chkDefaultPersonal" runat="server" AutoPostBack="true" Visible="false"/>--%>
        <asp:DropDownList ID="ddlLoc" CssClass="ddlLoc" runat="server" AutoPostBack="true" Visible="false">
            <asp:ListItem Value="0">-SELECT LINE-</asp:ListItem>
            <asp:ListItem Value="1">LINE 1</asp:ListItem>
            <asp:ListItem Value="2">LINE 2</asp:ListItem>
            <asp:ListItem Value="3">OFFLINE</asp:ListItem>
        </asp:DropDownList>
        <asp:Panel ID="pnlPalletHeader" CssClass="pnlPalletHeader" runat="server" Visible="false">
            <asp:Label ID="lblFilterLot" CssClass="lblFilterLot" runat="server" Text="Lot Num" Visible="True"></asp:Label>
            <asp:TextBox ID="txtFilterLot" CssClass="txtFilterLot" runat="server" Visible="True" Enabled="false"  AutoPostBack="true"></asp:TextBox>  
            <asp:Label ID="lblWeek" CssClass="lblWeek" runat="server" Text="Week" Visible="True" ></asp:Label>
            <asp:TextBox ID="txtWeek" CssClass="txtWeek" runat="server" Visible="True" Enabled="false" AutoPostBack="true" CausesValidation="false"></asp:TextBox>  
            <asp:RequiredFieldValidator ID="reqFieldValWeek" CssClass="reqFieldValWeek" runat="server" ErrorMessage="***" ControlToValidate="txtWeek" Enabled="False"></asp:RequiredFieldValidator>   
            <asp:CustomValidator ID="cstValWeek" CssClass="cstValWeek" runat="server" ErrorMessage="VALUE IS INVALID!" ControlToValidate="txtWeek" Enabled="False"></asp:CustomValidator>
            <asp:Label ID="lblYear" CssClass="lblYear" runat="server" Text="Year" Visible="True"></asp:Label>
            <asp:TextBox ID="txtYear" CssClass="txtYear" runat="server" Visible="True" Enabled="false"  AutoPostBack="true"></asp:TextBox>    
            <asp:RequiredFieldValidator ID="reqFieldValYear" CssClass="reqFieldValYear" runat="server" ErrorMessage="***" ControlToValidate="txtYear" Enabled="False"></asp:RequiredFieldValidator>   
            <asp:RangeValidator ID="rngValYear" CssClass="rngValYear" runat="server" ErrorMessage="MUST BE VALID YEAR 'XXXX'" ControlToValidate="txtYear" Enabled="False"
                MinimumValue="2010"
                MaximumValue="2021"
                Type="Integer"                  
                Text="MUST BE VALID YEAR 'XXXX'" />            
            <asp:Label ID="lblPallet" CssClass="lblPallet" runat="server" Text="Pallet" Visible="True"></asp:Label>
            <asp:TextBox ID="txtPallet" CssClass="txtPallet" runat="server" Visible="True" Enabled="false"  AutoPostBack="true"></asp:TextBox> 
            <asp:RequiredFieldValidator ID="reqFieldValPallet" CssClass="reqFieldValPallet" runat="server" ErrorMessage="***" ControlToValidate="txtPallet" Enabled="False"></asp:RequiredFieldValidator>                
            <asp:Button ID="btnSubmit" CssClass="btnSubmit" runat="server" Text="Submit" CausesValidation="True" Visible="true" Enabled="false" />
        </asp:Panel>
        <asp:Panel ID="pnlVerify" CssClass="pnlVerify" runat="server" Visible="false">
            <asp:Label ID="lblVerify" CssClass="lblVerify" runat="server" Text="Week is not current.  Accept?" ></asp:Label>
            <asp:Button ID="btnYes" CssClass="btnYes" runat="server" Text="Yes" /> 
            <asp:Button ID="btnNo" CssClass="btnNo" runat="server" Text="No" /> 
        </asp:Panel>
        <asp:Panel ID="pnlVerifyDelete" CssClass="pnlVerify" runat="server" Visible="false">
            <asp:Label ID="lblVerifyDelete" CssClass="lblVerify" runat="server" Text="Delete All?" ></asp:Label>
            <asp:Button ID="btnYesDelete" CssClass="btnYes" runat="server" Text="Yes" /> 
            <asp:Button ID="btnNoDelete" CssClass="btnNo" runat="server" Text="No" /> 
        </asp:Panel>
        <asp:Panel ID="pnlSerialData" CssClass="pnlSerialData" runat="server" Visible="False" BorderStyle="Solid" DefaultButton="btnFinish">
            <asp:Label ID="lblMode" CssClass="lblMode" runat="server" Text=""></asp:Label>
            <asp:ListBox ID="lstSerials" CssClass="lstSerials" runat="server"></asp:ListBox>
            <asp:Label ID="lblSerial" CssClass="lblSerial" runat="server" Text="Enter Barcode"></asp:Label>
            <asp:TextBox ID="txtSerial" CssClass="txtSerial" runat="server" AutoPostBack="True"></asp:TextBox>            
            <asp:Label ID="lblSerialCheck" runat="server" Text="" Visible="false"></asp:Label>
            <asp:TextBox ID="txtSerialByScan" CssClass="txtSerial" runat="server" AutoPostBack="True" Visible="false"></asp:TextBox> 
            <asp:Label ID="lblCurrCountText" CssClass="lblCurrCountText" runat="server" Text="Current Count:"></asp:Label>
            <asp:Label ID="lblCurrCount" CssClass="lblCurrCount" runat="server" Text="0"></asp:Label> 
            <asp:Label ID="lblCurrRowCountText" CssClass="lblCurrRowCountText" runat="server" Text="Row Count:"></asp:Label>
            <asp:Label ID="lblCurrRowCount" CssClass="lblCurrRowCount" runat="server" Text="0"></asp:Label>           
            <asp:Button ID="btnDelete" CssClass="btnDelete" runat="server" Text="Delete Serial" Enabled="False"/>
            <asp:Button ID="btnDeleteAll" CssClass="btnDeleteAll" runat="server" Text="Delete All" Enabled="False"/>
            <asp:Button ID="btnDeleteByScan" CssClass="btnDeleteByScan" runat="server" Text="Delete Serial By Scan" Enabled="False"/>
            <asp:Button ID="btnFinish" CssClass="btnFinish" runat="server" Text="Finish Pallet" CausesValidation="True" />
            <asp:Label ID="lblCompletedOn" CssClass="lblCompletedOn" runat="server" Text="" Visible="false"></asp:Label>

            
            <asp:Panel ID="pnlSignOff" CssClass="pnlSignOff" runat="server" Height="150px" Width="150px" BorderStyle="solid" BorderWidth="1px" Visible="False">
                <asp:Label ID="lblInitials" CssClass="lblInitials" runat="server" Text="Signoff Initials"></asp:Label> 
                <asp:TextBox ID="txtInitials" CssClass="txtInitials" runat="server"></asp:TextBox> 
                <asp:RequiredFieldValidator ID="reqFieldInitials" CssClass="reqFieldInitials" runat="server" ErrorMessage="***" ControlToValidate="txtInitials" Enabled="False"></asp:RequiredFieldValidator>   
                <asp:Label ID="lblCountVerify" CssClass="lblCountVerify" runat="server" Text="Verify Serial Count"></asp:Label>   
                <asp:TextBox ID="txtCountVerify" CssClass="txtCountVerify" runat="server"></asp:TextBox>
                <asp:CustomValidator ID="cstValCountVerify" CssClass="cstValCountVerify" runat="server" ControlToValidate="txtCountVerify" ErrorMessage="***" Enabled="True"></asp:CustomValidator>
                <asp:RequiredFieldValidator ID="reqFieldCountVerify" CssClass="cstValCountVerify" runat="server" ErrorMessage="***" ControlToValidate="txtCountVerify" Enabled="False"></asp:RequiredFieldValidator>   
                <asp:Label ID="lblFinish" CssClass="lblFinish" runat="server" Text="Click Finish Again to Complete."></asp:Label>

            </asp:Panel>
        </asp:Panel>  
        <asp:GridView ID="grdDupBarcodes" CssClass="grdDupBarcodes" runat="server" AutoGenerateColumns="false" Visible="false">
            <HeaderStyle CssClass="grdDupBarcodesHeader"></HeaderStyle>        
                <RowStyle CssClass="grdDupBarcodesRow" /> 
                  <Columns>
                      <asp:BoundField DataField="BARCODE" HeaderText="BARCODE" ReadOnly="True" Visible="True" ItemStyle-Width="100" ItemStyle-HorizontalAlign="Center"  />
                      <asp:BoundField DataField="BOX_POSITION" HeaderText="BOX POSITION" ReadOnly="True" Visible="True" ItemStyle-Width="50" ItemStyle-HorizontalAlign="Center" />
                      <asp:BoundField DataField="PALLET_NUM" HeaderText="PALLET" ReadOnly="True" Visible="True" ItemStyle-Width="50" ItemStyle-HorizontalAlign="Center" />
                  </Columns>
            </asp:GridView>
        <asp:Panel ID="pnlDuplicateError" CssClass="pnlDuplicateError" runat="server" Visible="false">
            <asp:Button ID="btnDupOK" CssClass="btnDupOK" runat="server" Text="OK" />
        </asp:Panel>
    </div>
    </form>
</body>
</html>
