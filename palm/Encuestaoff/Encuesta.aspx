<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Encuesta.aspx.cs" Inherits="Encuesta" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <script type="text/javascript">
    function isNumberKey(evt) {
        var charCode = (evt.which) ? evt.which : evt.keyCode;
        if (charCode != 46 && charCode > 31
          && (charCode < 48 || charCode > 57))
            return false;

        return true;
    }

    function handleChange(input) {
        if (input.value < 0) input.value = 0;
        if (input.value > 100) input.value = 100;
    }
    //-->
    </script>
    
</head>
<body>
    <form id="form1" runat="server" style="width:1000px">
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
        </div>

        <br />

        <asp:Table ID="Table1" runat="server" HorizontalAlign="Center" BorderStyle="Solid" BorderWidth="1" CellSpacing="10">
            <asp:TableRow Style="margin-top:30px">
                <asp:TableCell>
                    Cadena:
                </asp:TableCell>
                <asp:TableCell Width="310px">
                    <asp:DropDownList ID="ddcadena" runat="server" OnSelectedIndexChanged="ddcadena_SelectedIndexChanged" AutoPostBack="true" Width="310px"></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow>
                <asp:TableCell>
                    Comuna:
                </asp:TableCell>
                <asp:TableCell>
                    <asp:DropDownList ID="ddcomuna" runat="server" OnSelectedIndexChanged="ddcomuna_SelectedIndexChanged" AutoPostBack="true" Width="310px"></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow>
                <asp:TableCell>
                    Pdv:
                </asp:TableCell>
                <asp:TableCell>
                    <asp:DropDownList ID="ddpdv" runat="server" Width="310px"></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>

        <br />

        <div style="width:1000px; text-align:center">
          <asp:Button ID="Inicio" runat="server" Text="Comenzar" OnClick="Inicio_Click"/>
        </div>

        <br />

        <div style="width:1000px; text-align:center">
          <asp:DropDownList ID="ddcategorias" runat="server" Width="310px" AutoPostBack="true" OnSelectedIndexChanged="ddcategorias_SelectedIndexChanged" Visible="false"></asp:DropDownList>
        </div>

        <asp:Panel id="Item1" style="width:1000px; text-align:center" runat="server" Visible="false">
            <br />
            <asp:Label ID="li1versus" runat="server" Text="" Visible="false"></asp:Label>
            <br />
            <asp:Table ID="Table2" runat="server" HorizontalAlign="Center" BorderStyle="Solid" BorderWidth="1">
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Width="10px"/>
                    <asp:TableCell Width="100px"/>
                    <asp:TableCell Width="110px"><asp:Label Width="110px" ID="li1titulodesa" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li1titulocom1" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li1titulocom2" runat="server"></asp:Label></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="1.-"/>
                    <asp:TableCell Text="Precio" HorizontalAlign="Left" />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li1preciodesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li1preciocom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li1preciocom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="2.-"/>
                    <asp:TableCell Text="Facing" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li1facingdesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li1facingcom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li1facingcom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="3.-"/>
                    <asp:TableCell Text="Disponibilidad" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1disdesa" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1discom1" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1discom2" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="4.-"/>
                    <asp:TableCell Text="Promoción" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1promdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1promcom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1promcom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="5.-"/>
                    <asp:TableCell Text="Ex. Adicional" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1exdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1excom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi1excom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>

        <asp:Panel id="Item2" style="width:1000px; text-align:center" runat="server" Visible="false">
            <br />
            <asp:Label ID="li2versus" runat="server" Text="" Visible="false"></asp:Label>
            <br />
            <asp:Table ID="Table3" runat="server" HorizontalAlign="Center" BorderStyle="Solid" BorderWidth="1">
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Width="10px"/>
                    <asp:TableCell Width="100px"/>
                    <asp:TableCell Width="110px"><asp:Label Width="110px" ID="li2titulodesa" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li2titulocom1" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li2titulocom2" runat="server"></asp:Label></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="1.-"/>
                    <asp:TableCell Text="Precio" HorizontalAlign="Left" />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li2preciodesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li2preciocom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li2preciocom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="2.-"/>
                    <asp:TableCell Text="Facing" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li2facingdesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li2facingcom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li2facingcom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="3.-"/>
                    <asp:TableCell Text="Disponibilidad" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2disdesa" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2discom1" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2discom2" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="4.-"/>
                    <asp:TableCell Text="Promoción" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2promdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2promcom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2promcom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="5.-"/>
                    <asp:TableCell Text="Ex. Adicional" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2exdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2excom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi2excom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>

        <asp:Panel id="Item3" style="width:1000px; text-align:center" runat="server" Visible="false">
            <br />
            <asp:Label ID="li3versus" runat="server" Text="" Visible="false"></asp:Label>
            <br />
            <asp:Table ID="Table4" runat="server" HorizontalAlign="Center" BorderStyle="Solid" BorderWidth="1">
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Width="10px"/>
                    <asp:TableCell Width="100px"/>
                    <asp:TableCell Width="110px"><asp:Label Width="110px" ID="li3titulodesa" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li3titulocom1" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li3titulocom2" runat="server"></asp:Label></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="1.-"/>
                    <asp:TableCell Text="Precio" HorizontalAlign="Left" />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li3preciodesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li3preciocom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li3preciocom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="2.-"/>
                    <asp:TableCell Text="Facing" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li3facingdesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li3facingcom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li3facingcom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="3.-"/>
                    <asp:TableCell Text="Disponibilidad" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3disdesa" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3discom1" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3discom2" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="4.-"/>
                    <asp:TableCell Text="Promoción" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3promdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3promcom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3promcom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="5.-"/>
                    <asp:TableCell Text="Ex. Adicional" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3exdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3excom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi3excom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>

        <asp:Panel id="Item4" style="width:1000px; text-align:center" runat="server" Visible="false">
            <br />
            <asp:Label ID="li4versus" runat="server" Text="" Visible="false"></asp:Label>
            <br />
            <asp:Table ID="Table5" runat="server" HorizontalAlign="Center" BorderStyle="Solid" BorderWidth="1">
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Width="10px"/>
                    <asp:TableCell Width="100px"/>
                    <asp:TableCell Width="110px"><asp:Label Width="110px" ID="li4titulodesa" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li4titulocom1" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li4titulocom2" runat="server"></asp:Label></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="1.-"/>
                    <asp:TableCell Text="Precio" HorizontalAlign="Left" />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li4preciodesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li4preciocom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li4preciocom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="2.-"/>
                    <asp:TableCell Text="Facing" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li4facingdesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li4facingcom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li4facingcom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="3.-"/>
                    <asp:TableCell Text="Disponibilidad" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4disdesa" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4discom1" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4discom2" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="4.-"/>
                    <asp:TableCell Text="Promoción" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4promdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4promcom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4promcom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="5.-"/>
                    <asp:TableCell Text="Ex. Adicional" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4exdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4excom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi4excom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>

        <asp:Panel id="Item5" style="width:1000px; text-align:center" runat="server" Visible="false">
            <br />
            <asp:Label ID="li5versus" runat="server" Text="" Visible="false"></asp:Label>
            <br />
            <asp:Table ID="Table6" runat="server" HorizontalAlign="Center" BorderStyle="Solid" BorderWidth="1">
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Width="10px"/>
                    <asp:TableCell Width="100px"/>
                    <asp:TableCell Width="110px"><asp:Label Width="110px" ID="li5titulodesa" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li5titulocom1" runat="server"></asp:Label></asp:TableCell>
                    <asp:TableCell Width="100px"><asp:Label Width="100px" ID="li5titulocom2" runat="server"></asp:Label></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="1.-"/>
                    <asp:TableCell Text="Precio" HorizontalAlign="Left" />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li5preciodesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li5preciocom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li5preciocom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="2.-"/>
                    <asp:TableCell Text="Facing" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:TextBox Width="100px" ID="li5facingdesa" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li5facingcom1" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                    <asp:TableCell><asp:TextBox Width="100px" ID="li5facingcom2" runat="server" onkeypress="return isNumberKey(event)" onFocus="this.select()" Text="0" Style="text-align:center" MaxLength="3" onchange="handleChange(this);"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="3.-"/>
                    <asp:TableCell Text="Disponibilidad" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5disdesa" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5discom1" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5discom2" runat="server"><asp:ListItem Text="Si" Value="S"/><asp:ListItem Text="No" Value="N"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="4.-"/>
                    <asp:TableCell Text="Promoción" HorizontalAlign="Left"  />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5promdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5promcom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5promcom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Style="margin-top:30px">
                    <asp:TableCell Text="5.-"/>
                    <asp:TableCell Text="Ex. Adicional" HorizontalAlign="Left" />
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5exdesa" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5excom1" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                    <asp:TableCell><asp:DropDownList Width="100px" ID="ddi5excom2" runat="server"><asp:ListItem Text="No" Value="N"/><asp:ListItem Text="Si" Value="S"/></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>

        <br />

        <div style="width:1000px; text-align:center">
            <asp:Button ID="Enviar" runat="server" Text="Enviar Respuestas" OnClick="Enviar_Click" Visible="false"/>
            <asp:Button ID="Npdv" runat="server" Text="Nuevo Pdv" OnClick="Npdv_Click" Visible="false" Style="margin-left:50px"/>
        </div>

    </form>
</body>
</html>
