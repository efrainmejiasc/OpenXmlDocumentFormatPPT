﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="OpenXmlDocumentFormatPPT._default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
      <div  style="margin-top:10%;">
        <asp:Label ID="Label2" runat="server" Text="Label"></asp:Label><br /><br />
        <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
    </div>
        <br /> <br /><br />
  
         <div>
                <asp:Button ID="btnEdit" runat="server" Text="Editar" OnClick="btnEdit_Click" />
               <asp:Button ID="btnDownload" runat="server" Text="Descargar" OnClick="btnDownload_Click" />
        </div>
    </form>
</body>
</html>
