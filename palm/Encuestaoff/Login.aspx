<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Login.aspx.cs" Inherits="Login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        .auto-style1 {
            height: 21px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table>
                <tr>
                    <td style="width: 20%; height: 84px;">
                        <a href="#">
                        <img src="Imagenes/LogoFinal.jpg" style="border-width: 0px; width: 218px; height: 49px" />
                        </a>
                    </td>
                    <td style="width: 20%; height: 84px;">&nbsp;</td>
                    <td style="width: 20%; height: 84px;">&nbsp;</td>
                    <td style="width: 20%; height: 84px;">&nbsp;</td>
                    <td style="width: 21%; height: 84px;">&nbsp;</td>
                </tr>
                <tr style="background-color: #747B5A" class="nuevoEstilo1">
                    <td colspan="5" style="width: 20%; height: 10px; background-color: #e4e0c5;">&nbsp;</td>
                </tr>
            </table>
            
            <br />
            <br />
            <br />

            <table cellpadding="0" frame="border" style="margin-left:25%; width:350px">
                <tr>
                    <td align="center" colspan="2" style="color:White;background-color:#6B696B;font-weight:bold;" class="auto-style1">Iniciar sesión</td>
                </tr>
                <tr>
                    <td align="right" style="padding-top:20px">
                        <asp:Label ID="UserNameLabel" runat="server">Nombre de usuario:</asp:Label>
                    </td>
                    <td style="padding-top:20px">
                        <asp:DropDownList ID="ddusuarios" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="padding-top:20px" >
                        <asp:Label ID="PasswordLabel" runat="server">Contraseña:</asp:Label>
                    </td>
                    <td style="padding-top:20px">
                        <asp:TextBox ID="LPass" runat="server" TextMode="Password"></asp:TextBox>  
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="2">
                        <br />
                        <asp:Button ID="Acceder" runat="server" Text="Acceder" OnClick="Acceder_Click" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
