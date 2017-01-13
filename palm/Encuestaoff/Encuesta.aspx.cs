using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;

public partial class Encuesta : System.Web.UI.Page
{
    SqlConnection cnn = new SqlConnection("Data Source=SQLSERVER;Initial Catalog=DesaERP;Persist Security Info=True;User ID=sa;password=desakey");
    String SQL = String.Empty;
    SqlCommand cmd;
    SqlDataReader reader;
    String idcadena = String.Empty;
    String idcomuna = String.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Session["idusuario"] == null)
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Error", "alert('Favor... Inicie Sesión...');window.location ='Login.aspx';", true);
            }
            else
            {
                CargarCadena();
            }
        }
    }

    //Carga Cadena del filtro
    protected void CargarCadena()
    {
        ddcadena.Items.Clear();

        cnn.Open();

        SQL = " SELECT DISTINCT T1.[idcadena],T1.[nombre] " +
              " FROM [DataWarehouse].[dbo].[ENC_PDV] AS T0 " +
              " INNER JOIN [DataWarehouse].[dbo].[ENC_CADENA] AS T1 " +
              " ON T0.idcadena=T1.idcadena " +
              " Order by T1.nombre ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        ddcadena.Items.Add(new ListItem("Todas", ""));

        while (reader.Read())
        {
            ddcadena.Items.Add(new ListItem(reader.GetString(1),reader.GetInt32(0).ToString()));
        }

        reader.Close();
        cnn.Close();

        //Llama a la carga de Comunas
        CargarComunas();
    }

    //Carga de Comunas en el Filtro
    protected void CargarComunas()
    {
        ddcomuna.Items.Clear();

        idcadena = ddcadena.SelectedValue.ToString();

        cnn.Open();

        SQL = " SELECT DISTINCT T1.[idcomuna],T1.[nombre] " +
              " FROM [DataWarehouse].[dbo].[ENC_PDV] as T0 " +
              " INNER JOIN [DataWarehouse].[dbo].[ENC_COMUNAS] as T1 " +
              " ON T0.idcomuna=T1.idcomuna " +
              " Where T0.idcadena like '"+idcadena+"%' " +
              " Order by T1.nombre ";
        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        ddcomuna.Items.Add(new ListItem("Todas", ""));

        while (reader.Read())
        {
            ddcomuna.Items.Add(new ListItem(reader.GetString(1), reader.GetDecimal(0).ToString()));
        }

        reader.Close();
        cnn.Close();

        //Llama a la carga de PDV
        CargaPdv();
    }

    //Carga los PDV segun datos ingresados anteriormente
    protected void CargaPdv()
    {
        ddpdv.Items.Clear();

        idcadena = ddcadena.SelectedValue.ToString();
        idcomuna = ddcomuna.SelectedValue.ToString();
        cnn.Open();

        SQL = " SELECT DISTINCT [idpdv],[nombrepdv] " +
              " FROM [DataWarehouse].[dbo].[ENC_PDV] AS T0 " +
              " Where idcadena like '"+idcadena+"%' and idcomuna like '"+idcomuna+"%' ";
        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        while (reader.Read())
        {
            ddpdv.Items.Add(new ListItem(reader.GetString(1), reader.GetString(0)));
        }

        reader.Close();
        cnn.Close();
    }

    //En caso de que cambie la cadena, se llama a la carga de comuna
    protected void ddcadena_SelectedIndexChanged(object sender, EventArgs e)
    {
        CargarComunas();
    }

    //En caso de cambiar la comuna, se llama la carga de los PDV
    protected void ddcomuna_SelectedIndexChanged(object sender, EventArgs e)
    {
        CargaPdv();
    }

    //Boton iniciar Encuesta
    protected void Inicio_Click(object sender, EventArgs e)
    {
        Session["idpdv"] = ddpdv.SelectedValue.ToString();
        Session["indmsg"] = "N";
        //Bloqueo de elementos de inicio
        Table1.Enabled = false;
        Inicio.Enabled = false;
        //Desbloqueo de elementos para generar la encuesta
        ddcategorias.Visible = true;
        Enviar.Visible = true;
        Npdv.Visible = true;

        //Carga de Categoria
        CargaCategorias();
    }

    //Se cargan las categorias de la encuesta
    protected void CargaCategorias()
    {
        ddcategorias.Items.Clear();

        cnn.Open();

        SQL = " SELECT [idcategoria],[nombre] " +
              " FROM [DataWarehouse].[dbo].[ENC_CATEGORIA] ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        while (reader.Read())
        {
            ddcategorias.Items.Add(new ListItem(reader.GetString(1), reader.GetInt32(0).ToString()));
        }

        reader.Close();
        cnn.Close();

        //Corrobora si existen respuestas
        if (Existe())
        {
            CargaDatos();
        }
        else
        {
            //Se solicita la carga de Items
            CargaItems();
        }

        Session["before"] = ddcategorias.SelectedValue.ToString();
    }

    //Si se cambia la categoria, se recarga el layout y los items
    protected void ddcategorias_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (Validar())
        {
            if (Before())
            {
                Session["indmsg"] = "Y";
                Enviar_Click(sender, e);
                Session["indmsg"] = "N";
            }

            //Corrobora si existen respuestas
            if (Existe())
            {
                RestartItem();
                CargaDatos();
            }
            else
            {
                //Se solicita la carga de Items
                RestartItem();
                CargaItems();
            }

            Session["before"] = ddcategorias.SelectedValue.ToString();
        }
        else
        {
            ddcategorias.SelectedIndex=ddcategorias.Items.IndexOf(ddcategorias.Items.FindByValue(Session["before"].ToString()));
        }
    }

    //Se verifica si existen respuestas, para el local y el dia.
    protected Boolean Before()
    {
        String idcategoria = Session["before"].ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();
        Boolean aux1 = false;

        Boolean pdesa = false;
        Boolean pcom1 = false;
        Boolean pcom2 = false;
        Boolean fdesa = false;
        Boolean fcom1 = false;
        Boolean fcom2 = false;
        Boolean ddesa = false;
        Boolean dcom1 = false;
        Boolean dcom2 = false;
        Boolean pmdesa = false;
        Boolean pmcom1 = false;
        Boolean pmcom2 = false;
        Boolean edesa = false;
        Boolean ecom1 = false;
        Boolean ecom2 = false;

        cnn.Open();

        SQL = " SELECT T1.[iditem], T1.[nombredesa], T1.[nombrecomp1], T1.[nombrecomp2], " +
              " T0.[preciodesa], T0.[preciocomp1], T0.[preciocomp2], " +
              " T0.[facingdesa], T0.[facingcomp1], T0.[facingcomp2], " +
              " T0.[dispdesa], T0.[dispcomp1], T0.[dispcomp2], " +
              " T0.[promdesa], T0.[promcomp1], T0.[promcomp2], " +
              " T0.[exbadicdesa], T0.[exbadiccomp1], T0.[exbadiccomp2] " +
              " FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] AS T0 " +
              " INNER JOIN [DataWarehouse].[dbo].[ENC_ITEM] AS T1 " +
              " ON T0.iditem=T1.iditem  " +
              " Where T1.idcategoria=" + idcategoria + " and T0.fecha='" + fecha + "' and T0.idpdv='" + idpdv + "'";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        // Compara Item 1
        if (reader.Read())
        {
            pdesa = reader.GetInt32(4).ToString().Equals(li1preciodesa.Text) ? false : true;
            pcom1 = reader.GetInt32(5).ToString().Equals(li1preciocom1.Text) ? false : true;
            pcom2 = reader.GetInt32(6).ToString().Equals(li1preciocom2.Text) ? false : true;
            fdesa = reader.GetInt32(7).ToString().Equals(li1facingdesa.Text) ? false : true;
            fcom1 = reader.GetInt32(8).ToString().Equals(li1facingcom1.Text) ? false : true;
            fcom2 = reader.GetInt32(9).ToString().Equals(li1facingcom2.Text) ? false : true;
            ddesa = reader.GetString(10).Equals(ddi1disdesa.SelectedValue.ToString()) ? false : true;
            dcom1 = reader.GetString(11).Equals(ddi1discom1.SelectedValue.ToString()) ? false : true;
            dcom2 = reader.GetString(12).Equals(ddi1discom2.SelectedValue.ToString()) ? false : true;
            pmdesa = reader.GetString(13).Equals(ddi1promdesa.SelectedValue.ToString()) ? false : true;
            pmcom1 = reader.GetString(14).Equals(ddi1promcom1.SelectedValue.ToString()) ? false : true;
            pmcom2 = reader.GetString(15).Equals(ddi1promcom2.SelectedValue.ToString()) ? false : true;
            edesa = reader.GetString(16).Equals(ddi1exdesa.SelectedValue.ToString()) ? false : true;
            ecom1 = reader.GetString(17).Equals(ddi1excom1.SelectedValue.ToString()) ? false : true;
            ecom2 = reader.GetString(18).Equals(ddi1excom2.SelectedValue.ToString()) ? false : true;

            if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
            {
                aux1 = true;
            }
        }
        else
        {
            if(!li1versus.Text.Equals(""))
            {
                pdesa = li1preciodesa.Text.Equals("0") ? false : true;
                pcom1 = li1preciocom1.Text.Equals("0") ? false : true;
                pcom2 = li1preciocom2.Text.Equals("0") ? false : true;
                fdesa = li1facingdesa.Text.Equals("0") ? false : true;
                fcom1 = li1facingcom1.Text.Equals("0") ? false : true;
                fcom2 = li1facingcom2.Text.Equals("0") ? false : true;
                ddesa = ddi1disdesa.SelectedValue.ToString().Equals("S") ? false : true;
                dcom1 = ddi1discom1.SelectedValue.ToString().Equals("S") ? false : true;
                dcom2 = ddi1discom2.SelectedValue.ToString().Equals("S") ? false : true;
                pmdesa = ddi1promdesa.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom1 = ddi1promcom1.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom2 = ddi1promcom2.SelectedValue.ToString().Equals("N") ? false : true;
                edesa = ddi1exdesa.SelectedValue.ToString().Equals("N") ? false : true;
                ecom1 = ddi1excom1.SelectedValue.ToString().Equals("N") ? false : true;
                ecom2 = ddi1excom2.SelectedValue.ToString().Equals("N") ? false : true;

                if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
                {
                    aux1 = true;
                }
            }
        }

        // Compara Item 2
        if (reader.Read())
        {
            pdesa = reader.GetInt32(4).ToString().Equals(li2preciodesa.Text) ? false : true;
            pcom1 = reader.GetInt32(5).ToString().Equals(li2preciocom1.Text) ? false : true;
            pcom2 = reader.GetInt32(6).ToString().Equals(li2preciocom2.Text) ? false : true;
            fdesa = reader.GetInt32(7).ToString().Equals(li2facingdesa.Text) ? false : true;
            fcom1 = reader.GetInt32(8).ToString().Equals(li2facingcom1.Text) ? false : true;
            fcom2 = reader.GetInt32(9).ToString().Equals(li2facingcom2.Text) ? false : true;
            ddesa = reader.GetString(10).Equals(ddi2disdesa.SelectedValue.ToString()) ? false : true;
            dcom1 = reader.GetString(11).Equals(ddi2discom1.SelectedValue.ToString()) ? false : true;
            dcom2 = reader.GetString(12).Equals(ddi2discom2.SelectedValue.ToString()) ? false : true;
            pmdesa = reader.GetString(13).Equals(ddi2promdesa.SelectedValue.ToString()) ? false : true;
            pmcom1 = reader.GetString(14).Equals(ddi2promcom1.SelectedValue.ToString()) ? false : true;
            pmcom2 = reader.GetString(15).Equals(ddi2promcom2.SelectedValue.ToString()) ? false : true;
            edesa = reader.GetString(16).Equals(ddi2exdesa.SelectedValue.ToString()) ? false : true;
            ecom1 = reader.GetString(17).Equals(ddi2excom1.SelectedValue.ToString()) ? false : true;
            ecom2 = reader.GetString(18).Equals(ddi2excom2.SelectedValue.ToString()) ? false : true;

            if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
            {
                aux1 = true;
            }
        }
        else
        {
            if (!li2versus.Text.Equals(""))
            {
                pdesa = li2preciodesa.Text.Equals("0") ? false : true;
                pcom1 = li2preciocom1.Text.Equals("0") ? false : true;
                pcom2 = li2preciocom2.Text.Equals("0") ? false : true;
                fdesa = li2facingdesa.Text.Equals("0") ? false : true;
                fcom1 = li2facingcom1.Text.Equals("0") ? false : true;
                fcom2 = li2facingcom2.Text.Equals("0") ? false : true;
                ddesa = ddi2disdesa.SelectedValue.ToString().Equals("S") ? false : true;
                dcom1 = ddi2discom1.SelectedValue.ToString().Equals("S") ? false : true;
                dcom2 = ddi2discom2.SelectedValue.ToString().Equals("S") ? false : true;
                pmdesa = ddi2promdesa.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom1 = ddi2promcom1.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom2 = ddi2promcom2.SelectedValue.ToString().Equals("N") ? false : true;
                edesa = ddi2exdesa.SelectedValue.ToString().Equals("N") ? false : true;
                ecom1 = ddi2excom1.SelectedValue.ToString().Equals("N") ? false : true;
                ecom2 = ddi2excom2.SelectedValue.ToString().Equals("N") ? false : true;

                if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
                {
                    aux1 = true;
                }
            }
        }

        // Compara Item 3
        if (reader.Read())
        {
            pdesa = reader.GetInt32(4).ToString().Equals(li3preciodesa.Text) ? false : true;
            pcom1 = reader.GetInt32(5).ToString().Equals(li3preciocom1.Text) ? false : true;
            pcom2 = reader.GetInt32(6).ToString().Equals(li3preciocom2.Text) ? false : true;
            fdesa = reader.GetInt32(7).ToString().Equals(li3facingdesa.Text) ? false : true;
            fcom1 = reader.GetInt32(8).ToString().Equals(li3facingcom1.Text) ? false : true;
            fcom2 = reader.GetInt32(9).ToString().Equals(li3facingcom2.Text) ? false : true;
            ddesa = reader.GetString(10).Equals(ddi3disdesa.SelectedValue.ToString()) ? false : true;
            dcom1 = reader.GetString(11).Equals(ddi3discom1.SelectedValue.ToString()) ? false : true;
            dcom2 = reader.GetString(12).Equals(ddi3discom2.SelectedValue.ToString()) ? false : true;
            pmdesa = reader.GetString(13).Equals(ddi3promdesa.SelectedValue.ToString()) ? false : true;
            pmcom1 = reader.GetString(14).Equals(ddi3promcom1.SelectedValue.ToString()) ? false : true;
            pmcom2 = reader.GetString(15).Equals(ddi3promcom2.SelectedValue.ToString()) ? false : true;
            edesa = reader.GetString(16).Equals(ddi3exdesa.SelectedValue.ToString()) ? false : true;
            ecom1 = reader.GetString(17).Equals(ddi3excom1.SelectedValue.ToString()) ? false : true;
            ecom2 = reader.GetString(18).Equals(ddi3excom2.SelectedValue.ToString()) ? false : true;

            if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
            {
                aux1 = true;
            }
        }
        else
        {
            if (!li3versus.Text.Equals(""))
            {
                pdesa = li3preciodesa.Text.Equals("0") ? false : true;
                pcom1 = li3preciocom1.Text.Equals("0") ? false : true;
                pcom2 = li3preciocom2.Text.Equals("0") ? false : true;
                fdesa = li3facingdesa.Text.Equals("0") ? false : true;
                fcom1 = li3facingcom1.Text.Equals("0") ? false : true;
                fcom2 = li3facingcom2.Text.Equals("0") ? false : true;
                ddesa = ddi3disdesa.SelectedValue.ToString().Equals("S") ? false : true;
                dcom1 = ddi3discom1.SelectedValue.ToString().Equals("S") ? false : true;
                dcom2 = ddi3discom2.SelectedValue.ToString().Equals("S") ? false : true;
                pmdesa = ddi3promdesa.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom1 = ddi3promcom1.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom2 = ddi3promcom2.SelectedValue.ToString().Equals("N") ? false : true;
                edesa = ddi3exdesa.SelectedValue.ToString().Equals("N") ? false : true;
                ecom1 = ddi3excom1.SelectedValue.ToString().Equals("N") ? false : true;
                ecom2 = ddi3excom2.SelectedValue.ToString().Equals("N") ? false : true;

                if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
                {
                    aux1 = true;
                }
            }
        }

        // Compara Item 4
        if (reader.Read())
        {
            pdesa = reader.GetInt32(4).ToString().Equals(li4preciodesa.Text) ? false : true;
            pcom1 = reader.GetInt32(5).ToString().Equals(li4preciocom1.Text) ? false : true;
            pcom2 = reader.GetInt32(6).ToString().Equals(li4preciocom2.Text) ? false : true;
            fdesa = reader.GetInt32(7).ToString().Equals(li4facingdesa.Text) ? false : true;
            fcom1 = reader.GetInt32(8).ToString().Equals(li4facingcom1.Text) ? false : true;
            fcom2 = reader.GetInt32(9).ToString().Equals(li4facingcom2.Text) ? false : true;
            ddesa = reader.GetString(10).Equals(ddi4disdesa.SelectedValue.ToString()) ? false : true;
            dcom1 = reader.GetString(11).Equals(ddi4discom1.SelectedValue.ToString()) ? false : true;
            dcom2 = reader.GetString(12).Equals(ddi4discom2.SelectedValue.ToString()) ? false : true;
            pmdesa = reader.GetString(13).Equals(ddi4promdesa.SelectedValue.ToString()) ? false : true;
            pmcom1 = reader.GetString(14).Equals(ddi4promcom1.SelectedValue.ToString()) ? false : true;
            pmcom2 = reader.GetString(15).Equals(ddi4promcom2.SelectedValue.ToString()) ? false : true;
            edesa = reader.GetString(16).Equals(ddi4exdesa.SelectedValue.ToString()) ? false : true;
            ecom1 = reader.GetString(17).Equals(ddi4excom1.SelectedValue.ToString()) ? false : true;
            ecom2 = reader.GetString(18).Equals(ddi4excom2.SelectedValue.ToString()) ? false : true;

            if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
            {
                aux1 = true;
            }
        }
        else
        {
            if (!li4versus.Text.Equals(""))
            {
                pdesa = li4preciodesa.Text.Equals("0") ? false : true;
                pcom1 = li4preciocom1.Text.Equals("0") ? false : true;
                pcom2 = li4preciocom2.Text.Equals("0") ? false : true;
                fdesa = li4facingdesa.Text.Equals("0") ? false : true;
                fcom1 = li4facingcom1.Text.Equals("0") ? false : true;
                fcom2 = li4facingcom2.Text.Equals("0") ? false : true;
                ddesa = ddi4disdesa.SelectedValue.ToString().Equals("S") ? false : true;
                dcom1 = ddi4discom1.SelectedValue.ToString().Equals("S") ? false : true;
                dcom2 = ddi4discom2.SelectedValue.ToString().Equals("S") ? false : true;
                pmdesa = ddi4promdesa.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom1 = ddi4promcom1.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom2 = ddi4promcom2.SelectedValue.ToString().Equals("N") ? false : true;
                edesa = ddi4exdesa.SelectedValue.ToString().Equals("N") ? false : true;
                ecom1 = ddi4excom1.SelectedValue.ToString().Equals("N") ? false : true;
                ecom2 = ddi4excom2.SelectedValue.ToString().Equals("N") ? false : true;

                if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
                {
                    aux1 = true;
                }
            }
        }

        // Compara Item 5
        if (reader.Read())
        {
            pdesa = reader.GetInt32(4).ToString().Equals(li5preciodesa.Text) ? false : true;
            pcom1 = reader.GetInt32(5).ToString().Equals(li5preciocom1.Text) ? false : true;
            pcom2 = reader.GetInt32(6).ToString().Equals(li5preciocom2.Text) ? false : true;
            fdesa = reader.GetInt32(7).ToString().Equals(li5facingdesa.Text) ? false : true;
            fcom1 = reader.GetInt32(8).ToString().Equals(li5facingcom1.Text) ? false : true;
            fcom2 = reader.GetInt32(9).ToString().Equals(li5facingcom2.Text) ? false : true;
            ddesa = reader.GetString(10).Equals(ddi5disdesa.SelectedValue.ToString()) ? false : true;
            dcom1 = reader.GetString(11).Equals(ddi5discom1.SelectedValue.ToString()) ? false : true;
            dcom2 = reader.GetString(12).Equals(ddi5discom2.SelectedValue.ToString()) ? false : true;
            pmdesa = reader.GetString(13).Equals(ddi5promdesa.SelectedValue.ToString()) ? false : true;
            pmcom1 = reader.GetString(14).Equals(ddi5promcom1.SelectedValue.ToString()) ? false : true;
            pmcom2 = reader.GetString(15).Equals(ddi5promcom2.SelectedValue.ToString()) ? false : true;
            edesa = reader.GetString(16).Equals(ddi5exdesa.SelectedValue.ToString()) ? false : true;
            ecom1 = reader.GetString(17).Equals(ddi5excom1.SelectedValue.ToString()) ? false : true;
            ecom2 = reader.GetString(18).Equals(ddi5excom2.SelectedValue.ToString()) ? false : true;

            if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
            {
                aux1 = true;
            }
        }
        else
        {
            if (!li5versus.Text.Equals(""))
            {
                pdesa = li5preciodesa.Text.Equals("0") ? false : true;
                pcom1 = li5preciocom1.Text.Equals("0") ? false : true;
                pcom2 = li5preciocom2.Text.Equals("0") ? false : true;
                fdesa = li5facingdesa.Text.Equals("0") ? false : true;
                fcom1 = li5facingcom1.Text.Equals("0") ? false : true;
                fcom2 = li5facingcom2.Text.Equals("0") ? false : true;
                ddesa = ddi5disdesa.SelectedValue.ToString().Equals("S") ? false : true;
                dcom1 = ddi5discom1.SelectedValue.ToString().Equals("S") ? false : true;
                dcom2 = ddi5discom2.SelectedValue.ToString().Equals("S") ? false : true;
                pmdesa = ddi5promdesa.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom1 = ddi5promcom1.SelectedValue.ToString().Equals("N") ? false : true;
                pmcom2 = ddi5promcom2.SelectedValue.ToString().Equals("N") ? false : true;
                edesa = ddi5exdesa.SelectedValue.ToString().Equals("N") ? false : true;
                ecom1 = ddi5excom1.SelectedValue.ToString().Equals("N") ? false : true;
                ecom2 = ddi5excom2.SelectedValue.ToString().Equals("N") ? false : true;

                if (pdesa || pcom1 || pcom2 || fdesa || fcom1 || fcom2 || ddesa || dcom1 || dcom2 || pmdesa || pmcom1 || pmcom2 || edesa || ecom1 || ecom2)
                {
                    aux1 = true;
                }
            }
        }

        reader.Close();
        cnn.Close();

        return aux1;
    }

    //Se verifica si existen respuestas, para el local y el dia.
    protected Boolean Existe()
    {
        String idcategoria = ddcategorias.SelectedValue.ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();
        Boolean aux1 = false;

        cnn.Open();

        SQL = " SELECT T0.[idrespuesta] " +
              " FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] AS T0 " +
              " INNER JOIN [DataWarehouse].[dbo].[ENC_ITEM] AS T1 " +
              " ON T0.iditem=T1.iditem Where T1.idcategoria='"+idcategoria+"' and T0.fecha="+fecha+" AND T0.idpdv='"+idpdv+"' ";


        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            Session["borrar"] = "T";
            aux1 = true;
        }
        else
        {
            Session["borrar"] = "F";
        }

        reader.Close();
        cnn.Close();

        return aux1;
    }

    //Se cargan los items a encuestar
    protected void CargaDatos()
    {
        String idcategoria = ddcategorias.SelectedValue.ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();

        cnn.Open();

        SQL = " SELECT T1.[iditem], T1.[nombredesa], T1.[nombrecomp1], T1.[nombrecomp2], " +
              " T0.[preciodesa], T0.[preciocomp1], T0.[preciocomp2], " +
              " T0.[facingdesa], T0.[facingcomp1], T0.[facingcomp2], " +
              " T0.[dispdesa], T0.[dispcomp1], T0.[dispcomp2], " +
              " T0.[promdesa], T0.[promcomp1], T0.[promcomp2], " +
              " T0.[exbadicdesa], T0.[exbadiccomp1], T0.[exbadiccomp2] " +
              " FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] AS T0 " +
              " INNER JOIN [DataWarehouse].[dbo].[ENC_ITEM] AS T1 " +
              " ON T0.iditem=T1.iditem  " +
              " Where T1.idcategoria="+idcategoria+" and T0.fecha='"+fecha+"' and T0.idpdv='"+idpdv+"'";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            li1versus.Text = reader.GetInt32(0).ToString();
            li1titulodesa.Text = reader.GetString(1);
            li1titulocom1.Text = reader.GetString(2);
            li1titulocom2.Text = reader.GetString(3);
            li1preciodesa.Text = reader.GetInt32(4).ToString();
            li1preciocom1.Text = reader.GetInt32(5).ToString();
            li1preciocom2.Text = reader.GetInt32(6).ToString();
            li1facingdesa.Text = reader.GetInt32(7).ToString();
            li1facingcom1.Text = reader.GetInt32(8).ToString();
            li1facingcom2.Text = reader.GetInt32(9).ToString();
            ddi1disdesa.SelectedIndex = reader.GetString(10).Equals("S") ? 0 : 1;
            ddi1discom1.SelectedIndex = reader.GetString(11).Equals("S") ? 0 : 1;
            ddi1discom2.SelectedIndex = reader.GetString(12).Equals("S") ? 0 : 1;
            ddi1promdesa .SelectedIndex = reader.GetString(13).Equals("S") ? 1 : 0;
            ddi1promcom1.SelectedIndex = reader.GetString(14).Equals("S") ? 1 : 0;
            ddi1promcom2.SelectedIndex = reader.GetString(15).Equals("S") ? 1 : 0;
            ddi1exdesa.SelectedIndex = reader.GetString(16).Equals("S") ? 1 : 0;
            ddi1excom1.SelectedIndex = reader.GetString(17).Equals("S") ? 1 : 0;
            ddi1excom2.SelectedIndex = reader.GetString(18).Equals("S") ? 1 : 0;
            Item1.Visible = true;
        }

        if (reader.Read())
        {
            li2versus.Text = reader.GetInt32(0).ToString();
            li2titulodesa.Text = reader.GetString(1);
            li2titulocom1.Text = reader.GetString(2);
            li2titulocom2.Text = reader.GetString(3);
            li2preciodesa.Text = reader.GetInt32(4).ToString();
            li2preciocom1.Text = reader.GetInt32(5).ToString();
            li2preciocom2.Text = reader.GetInt32(6).ToString();
            li2facingdesa.Text = reader.GetInt32(7).ToString();
            li2facingcom1.Text = reader.GetInt32(8).ToString();
            li2facingcom2.Text = reader.GetInt32(9).ToString();
            ddi2disdesa.SelectedIndex = reader.GetString(10).Equals("S") ? 0 : 1;
            ddi2discom1.SelectedIndex = reader.GetString(11).Equals("S") ? 0 : 1;
            ddi2discom2.SelectedIndex = reader.GetString(12).Equals("S") ? 0 : 1;
            ddi2promdesa.SelectedIndex = reader.GetString(13).Equals("S") ? 1 : 0;
            ddi2promcom1.SelectedIndex = reader.GetString(14).Equals("S") ? 1 : 0;
            ddi2promcom2.SelectedIndex = reader.GetString(15).Equals("S") ? 1 : 0;
            ddi2exdesa.SelectedIndex = reader.GetString(16).Equals("S") ? 1 : 0;
            ddi2excom1.SelectedIndex = reader.GetString(17).Equals("S") ? 1 : 0;
            ddi2excom2.SelectedIndex = reader.GetString(18).Equals("S") ? 1 : 0;
            Item2.Visible = true;
        }

        if (reader.Read())
        {
            li3versus.Text = reader.GetInt32(0).ToString();
            li3titulodesa.Text = reader.GetString(1);
            li3titulocom1.Text = reader.GetString(2);
            li3titulocom2.Text = reader.GetString(3);
            li3preciodesa.Text = reader.GetInt32(4).ToString();
            li3preciocom1.Text = reader.GetInt32(5).ToString();
            li3preciocom2.Text = reader.GetInt32(6).ToString();
            li3facingdesa.Text = reader.GetInt32(7).ToString();
            li3facingcom1.Text = reader.GetInt32(8).ToString();
            li3facingcom2.Text = reader.GetInt32(9).ToString();
            ddi3disdesa.SelectedIndex = reader.GetString(10).Equals("S") ? 0 : 1;
            ddi3discom1.SelectedIndex = reader.GetString(11).Equals("S") ? 0 : 1;
            ddi3discom2.SelectedIndex = reader.GetString(12).Equals("S") ? 0 : 1;
            ddi3promdesa.SelectedIndex = reader.GetString(13).Equals("S") ? 1 : 0;
            ddi3promcom1.SelectedIndex = reader.GetString(14).Equals("S") ? 1 : 0;
            ddi3promcom2.SelectedIndex = reader.GetString(15).Equals("S") ? 1 : 0;
            ddi3exdesa.SelectedIndex = reader.GetString(16).Equals("S") ? 1 : 0;
            ddi3excom1.SelectedIndex = reader.GetString(17).Equals("S") ? 1 : 0;
            ddi3excom2.SelectedIndex = reader.GetString(18).Equals("S") ? 1 : 0;
            Item3.Visible = true;
        }

        if (reader.Read())
        {
            li4versus.Text = reader.GetInt32(0).ToString();
            li4titulodesa.Text = reader.GetString(1);
            li4titulocom1.Text = reader.GetString(2);
            li4titulocom2.Text = reader.GetString(3);
            li4preciodesa.Text = reader.GetInt32(4).ToString();
            li4preciocom1.Text = reader.GetInt32(5).ToString();
            li4preciocom2.Text = reader.GetInt32(6).ToString();
            li4facingdesa.Text = reader.GetInt32(7).ToString();
            li4facingcom1.Text = reader.GetInt32(8).ToString();
            li4facingcom2.Text = reader.GetInt32(9).ToString();
            ddi4disdesa.SelectedIndex = reader.GetString(10).Equals("S") ? 0 : 1;
            ddi4discom1.SelectedIndex = reader.GetString(11).Equals("S") ? 0 : 1;
            ddi4discom2.SelectedIndex = reader.GetString(12).Equals("S") ? 0 : 1;
            ddi4promdesa.SelectedIndex = reader.GetString(13).Equals("S") ? 1 : 0;
            ddi4promcom1.SelectedIndex = reader.GetString(14).Equals("S") ? 1 : 0;
            ddi4promcom2.SelectedIndex = reader.GetString(15).Equals("S") ? 1 : 0;
            ddi4exdesa.SelectedIndex = reader.GetString(16).Equals("S") ? 1 : 0;
            ddi4excom1.SelectedIndex = reader.GetString(17).Equals("S") ? 1 : 0;
            ddi4excom2.SelectedIndex = reader.GetString(18).Equals("S") ? 1 : 0;
            Item4.Visible = true;
        }

        if (reader.Read())
        {
            li5versus.Text = reader.GetInt32(0).ToString();
            li5titulodesa.Text = reader.GetString(1);
            li5titulocom1.Text = reader.GetString(2);
            li5titulocom2.Text = reader.GetString(3);
            li5preciodesa.Text = reader.GetInt32(4).ToString();
            li5preciocom1.Text = reader.GetInt32(5).ToString();
            li5preciocom2.Text = reader.GetInt32(6).ToString();
            li5facingdesa.Text = reader.GetInt32(7).ToString();
            li5facingcom1.Text = reader.GetInt32(8).ToString();
            li5facingcom2.Text = reader.GetInt32(9).ToString();
            ddi5disdesa.SelectedIndex = reader.GetString(10).Equals("S") ? 0 : 1;
            ddi5discom1.SelectedIndex = reader.GetString(11).Equals("S") ? 0 : 1;
            ddi5discom2.SelectedIndex = reader.GetString(12).Equals("S") ? 0 : 1;
            ddi5promdesa.SelectedIndex = reader.GetString(13).Equals("S") ? 1 : 0;
            ddi5promcom1.SelectedIndex = reader.GetString(14).Equals("S") ? 1 : 0;
            ddi5promcom2.SelectedIndex = reader.GetString(15).Equals("S") ? 1 : 0;
            ddi5exdesa.SelectedIndex = reader.GetString(16).Equals("S") ? 1 : 0;
            ddi5excom1.SelectedIndex = reader.GetString(17).Equals("S") ? 1 : 0;
            ddi5excom2.SelectedIndex = reader.GetString(18).Equals("S") ? 1 : 0;
            Item5.Visible = true;
        }

        reader.Close();
        cnn.Close();
    }

    //Se cargan los items a encuestar
    protected void CargaItems()
    {
        String idcategoria = ddcategorias.SelectedValue.ToString();

        cnn.Open();

        SQL = " SELECT [iditem],[nombredesa],[nombrecomp1],[nombrecomp2] " +
              " FROM [DataWarehouse].[dbo].[ENC_ITEM] " +
              " Where idcategoria='"+idcategoria+"' " +
              " Order by iditem ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            li1versus.Text = reader.GetInt32(0).ToString();
            li1titulodesa.Text = reader.GetString(1);
            li1titulocom1.Text = reader.GetString(2);
            li1titulocom2.Text = reader.GetString(3);
            Item1.Visible = true;
        }

        if (reader.Read())
        {
            li2versus.Text = reader.GetInt32(0).ToString();
            li2titulodesa.Text = reader.GetString(1);
            li2titulocom1.Text = reader.GetString(2);
            li2titulocom2.Text = reader.GetString(3);
            Item2.Visible = true;
        }

        if (reader.Read())
        {
            li3versus.Text = reader.GetInt32(0).ToString();
            li3titulodesa.Text = reader.GetString(1);
            li3titulocom1.Text = reader.GetString(2);
            li3titulocom2.Text = reader.GetString(3);
            Item3.Visible = true;
        }

        if (reader.Read())
        {
            li4versus.Text = reader.GetInt32(0).ToString();
            li4titulodesa.Text = reader.GetString(1);
            li4titulocom1.Text = reader.GetString(2);
            li4titulocom2.Text = reader.GetString(3);
            Item4.Visible = true;
        }

        if (reader.Read())
        {
            li5versus.Text = reader.GetInt32(0).ToString();
            li5titulodesa.Text = reader.GetString(1);
            li5titulocom1.Text = reader.GetString(2);
            li5titulocom2.Text = reader.GetString(3);
            Item5.Visible = true;
        }

        reader.Close();
        cnn.Close();
    }

    //Se Reinician todos los elementos graficos del sitio
    protected void RestartItem()
    {
        li1versus.Text = "";
        li1titulodesa.Text = "";
        li1preciodesa.Text = "0";
        li1facingdesa.Text = "0";
        ddi1disdesa.SelectedValue = "S";
        ddi1promdesa.SelectedIndex = 0;
        ddi1exdesa.SelectedIndex = 0;
        li1titulocom1.Text = "";
        li1preciocom1.Text = "0";
        li1facingcom1.Text = "0";
        ddi1discom1.SelectedValue = "S";
        ddi1promcom1.SelectedIndex = 0;
        ddi1excom1.SelectedIndex = 0;
        li1titulocom2.Text = "";
        li1preciocom2.Text = "0";
        li1facingcom2.Text = "0";
        ddi1discom2.SelectedValue = "S";
        ddi1promcom2.SelectedIndex = 0;
        ddi1excom2.SelectedIndex = 0;
        Item1.Visible = false;

        li2versus.Text = "";
        li2titulodesa.Text = "";
        li2preciodesa.Text = "0";
        li2facingdesa.Text = "0";
        ddi2disdesa.SelectedValue = "S";
        ddi2promdesa.SelectedIndex = 0;
        ddi2exdesa.SelectedIndex = 0;
        li2titulocom1.Text = "";
        li2preciocom1.Text = "0";
        li2facingcom1.Text = "0";
        ddi2discom1.SelectedValue = "S";
        ddi2promcom1.SelectedIndex = 0;
        ddi2excom1.SelectedIndex = 0;
        li2titulocom2.Text = "";
        li2preciocom2.Text = "0";
        li2facingcom2.Text = "0";
        ddi2discom2.SelectedValue = "S";
        ddi2promcom2.SelectedIndex = 0;
        ddi2excom2.SelectedIndex = 0;
        Item2.Visible = false;

        li3versus.Text = "";
        li3titulodesa.Text = "";
        li3preciodesa.Text = "0";
        li3facingdesa.Text = "0";
        ddi3disdesa.SelectedValue = "S";
        ddi3promdesa.SelectedIndex = 0;
        ddi3exdesa.SelectedIndex = 0;
        li3titulocom1.Text = "";
        li3preciocom1.Text = "0";
        li3facingcom1.Text = "0";
        ddi3discom1.SelectedValue = "S";
        ddi3promcom1.SelectedIndex = 0;
        ddi3excom1.SelectedIndex = 0;
        li3titulocom2.Text = "";
        li3preciocom2.Text = "0";
        li3facingcom2.Text = "0";
        ddi3discom2.SelectedValue = "S";
        ddi3promcom2.SelectedIndex = 0;
        ddi3excom2.SelectedIndex = 0;
        Item3.Visible = false;

        li4versus.Text = "";
        li4titulodesa.Text = "";
        li4preciodesa.Text = "0";
        li4facingdesa.Text = "0";
        ddi4disdesa.SelectedValue = "S";
        ddi4promdesa.SelectedIndex = 0;
        ddi4exdesa.SelectedIndex = 0;
        li4titulocom1.Text = "";
        li4preciocom1.Text = "0";
        li4facingcom1.Text = "0";
        ddi4discom1.SelectedValue = "S";
        ddi4promcom1.SelectedIndex = 0;
        ddi4excom1.SelectedIndex = 0;
        li4titulocom2.Text = "";
        li4preciocom2.Text = "0";
        li4facingcom2.Text = "0";
        ddi4discom2.SelectedValue = "S";
        ddi4promcom2.SelectedIndex = 0;
        ddi4excom2.SelectedIndex = 0;
        Item4.Visible = false;

        li5versus.Text = "";
        li5titulodesa.Text = "";
        li5preciodesa.Text = "0";
        li5facingdesa.Text = "0";
        ddi5disdesa.SelectedValue = "S";
        ddi5promdesa.SelectedIndex = 0;
        ddi5exdesa.SelectedIndex = 0;
        li5titulocom1.Text = "";
        li5preciocom1.Text = "0";
        li5facingcom1.Text = "0";
        ddi5discom1.SelectedValue = "S";
        ddi5promcom1.SelectedIndex = 0;
        ddi5excom1.SelectedIndex = 0;
        li5titulocom2.Text = "";
        li5preciocom2.Text = "0";
        li5facingcom2.Text = "0";
        ddi5discom2.SelectedValue = "S";
        ddi5promcom2.SelectedIndex = 0;
        ddi5excom2.SelectedIndex = 0;
        Item5.Visible = false;
    }

    //Se genera el envio de las encuestas.
    protected void Enviar_Click(object sender, EventArgs e)
    {
        Boolean item1 = true;
        Boolean item2 = true;
        Boolean item3 = true;
        Boolean item4 = true;
        Boolean item5 = true;
        String borrar = Session["borrar"].ToString();

        //Se llama a la validacion de los datos.
        if (Validar())
        {
            if (!li1versus.Text.Equals(""))
            {
                if (borrar.Equals("T"))
                {
                    Borraritem1();
                }

                item1 = Insertaitem1();
            }

            if (!li2versus.Text.Equals(""))
            {
                if (borrar.Equals("T"))
                {
                    Borraritem2();
                }

                item2 = Insertaitem2(); 
            }

            if (!li3versus.Text.Equals(""))
            {
                if (borrar.Equals("T"))
                {
                    Borraritem3();
                }

                item3 = Insertaitem3(); 
            }

            if (!li4versus.Text.Equals(""))
            {
                if (borrar.Equals("T"))
                {
                    Borraritem4();
                }

                item4 = Insertaitem4(); 
            }

            if (!li5versus.Text.Equals(""))
            {
                if (borrar.Equals("T"))
                {
                    Borraritem5();
                }

                item5 = Insertaitem5(); 
            }

            if (Session["indmsg"].ToString().Equals("N"))
            {

                if (item1 == true && item2 == true && item3 == true && item4 == true && item5 == true)
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "Alert", "alert('Encuesta ingresada de manera correcta');", true);
                }
                else
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "Error", "alert('Error al ingresar la encuesta...');", true);
                }
            }
        }
    }

    //Se borra la informacion del item 1
    protected void Borraritem1()
    {
        String iditem = li1versus.Text;
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();

        cnn.Open();

        SQL = " DELETE FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
              " Where fecha='" + fecha + "' and idpdv='" + idpdv + "' and iditem='" + iditem + "'";

        cmd = new SqlCommand(SQL, cnn);
        cmd.ExecuteNonQuery();

        cnn.Close();
    }

    //Se borra la informacion del item 2
    protected void Borraritem2()
    {
        String iditem = li2versus.Text;
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();

        cnn.Open();

        SQL = " DELETE FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
              " Where fecha='" + fecha + "' and idpdv='" + idpdv + "' and iditem='" + iditem + "'";
        cmd = new SqlCommand(SQL, cnn);
        cmd.ExecuteNonQuery();

        cnn.Close();
    }

    //Se borra la informacion del item 3
    protected void Borraritem3()
    {
        String iditem = li3versus.Text;
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();

        cnn.Open();

        SQL = " DELETE FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
              " Where fecha='" + fecha + "' and idpdv='" + idpdv + "' and iditem='" + iditem + "'";

        cmd = new SqlCommand(SQL, cnn);
        cmd.ExecuteNonQuery();

        cnn.Close();
    }

    //Se borra la informacion del item 4
    protected void Borraritem4()
    {
        String iditem = li4versus.Text;
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();

        cnn.Open();

        SQL = " DELETE FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
              " Where fecha='" + fecha + "' and idpdv='" + idpdv + "' and iditem='" + iditem + "'";

        cmd = new SqlCommand(SQL, cnn);
        cmd.ExecuteNonQuery();

        cnn.Close();
    }

    //Se borra la informacion del item 5
    protected void Borraritem5()
    {
        String iditem = li5versus.Text;
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String idpdv = Session["idpdv"].ToString();

        cnn.Open();

        SQL = " DELETE FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
              " Where fecha='" + fecha + "' and idpdv='" + idpdv + "' and iditem='" + iditem + "'";

        cmd = new SqlCommand(SQL, cnn);
        cmd.ExecuteNonQuery();

        cnn.Close();
    }

    //Se genera el Ingreso de la primera encuesta si corresponde...
    protected Boolean Insertaitem1()
    {
        String idencuestador = Session["idusuario"].ToString();
        String idpdv = Session["idpdv"].ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String hora = DateTime.Now.ToString("HHmmss");
        String iditem = li1versus.Text;
        String preciodesa = li1preciodesa.Text;
        String preciocom1 = li1preciocom1.Text;
        String preciocom2 = li1preciocom2.Text;
        String facingdesa = li1facingdesa.Text;
        String facingcom1 = li1facingcom1.Text;
        String facingcom2 = li1facingcom2.Text;
        String dispdesa = ddi1disdesa.SelectedValue.ToString();
        String dispcom1 = ddi1discom1.SelectedValue.ToString();
        String dispcom2 = ddi1discom2.SelectedValue.ToString();
        String promdesa = ddi1promdesa.SelectedValue.ToString();
        String promcom1 = ddi1promcom1.SelectedValue.ToString();
        String promcom2 = ddi1promcom2.SelectedValue.ToString();
        String exdesa = ddi1exdesa.SelectedValue.ToString();
        String excom1 = ddi1excom1.SelectedValue.ToString();
        String excom2 = ddi1excom2.SelectedValue.ToString();
        Boolean validador = false;

        cnn.Open();

        SQL = " INSERT INTO [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
               " VALUES ('" + idencuestador + "', '" + idpdv + "', '" + fecha + "', '" + hora + "', '" + iditem + "','" + preciodesa + "', " +
               " '" + preciocom1 + "','" + preciocom2 + "','" + facingdesa + "','" + facingcom1 + "','" + facingcom2 + "','" + dispdesa + "','" + dispcom1 + "', " +
               " '" + dispcom2 + "','" + promdesa + "','" + promcom1 + "','" + promcom2 + "','" + exdesa + "','" + excom1 + "','" + excom2 + "');  " +
               " SELECT * FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] WHERE idrespuesta = SCOPE_IDENTITY(); ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            validador = true;
        }
        else
        {
            validador = false;
        }

        reader.Close();
        cnn.Close();

        return validador;
    }

    //Se genera el Ingreso de la segunda encuesta si corresponde...
    protected Boolean Insertaitem2()
    {
        String idencuestador = Session["idusuario"].ToString();
        String idpdv = Session["idpdv"].ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String hora = DateTime.Now.ToString("HHmmss");
        String iditem = li2versus.Text;
        String preciodesa = li2preciodesa.Text;
        String preciocom1 = li2preciocom1.Text;
        String preciocom2 = li2preciocom2.Text;
        String facingdesa = li2facingdesa.Text;
        String facingcom1 = li2facingcom1.Text;
        String facingcom2 = li2facingcom2.Text;
        String dispdesa = ddi2disdesa.SelectedValue.ToString();
        String dispcom1 = ddi2discom1.SelectedValue.ToString();
        String dispcom2 = ddi2discom2.SelectedValue.ToString();
        String promdesa = ddi2promdesa.SelectedValue.ToString();
        String promcom1 = ddi2promcom1.SelectedValue.ToString();
        String promcom2 = ddi2promcom2.SelectedValue.ToString();
        String exdesa = ddi2exdesa.SelectedValue.ToString();
        String excom1 = ddi2excom1.SelectedValue.ToString();
        String excom2 = ddi2excom2.SelectedValue.ToString();
        Boolean validador = false;

        cnn.Open();

        SQL = " INSERT INTO [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
               " VALUES ('" + idencuestador + "', '" + idpdv + "', '" + fecha + "', '" + hora + "', '" + iditem + "','" + preciodesa + "', " +
               " '" + preciocom1 + "','" + preciocom2 + "','" + facingdesa + "','" + facingcom1 + "','" + facingcom2 + "','" + dispdesa + "','" + dispcom1 + "', " +
               " '" + dispcom2 + "','" + promdesa + "','" + promcom1 + "','" + promcom2 + "','" + exdesa + "','" + excom1 + "','" + excom2 + "');  " +
               " SELECT * FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] WHERE idrespuesta = SCOPE_IDENTITY(); ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            validador = true;
        }
        else
        {
            validador = false;
        }

        reader.Close();
        cnn.Close();

        return validador;
    }

    //Se genera el Ingreso de la tercera encuesta si corresponde...
    protected Boolean Insertaitem3()
    {
        String idencuestador = Session["idusuario"].ToString();
        String idpdv = Session["idpdv"].ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String hora = DateTime.Now.ToString("HHmmss");
        String iditem = li3versus.Text;
        String preciodesa = li3preciodesa.Text;
        String preciocom1 = li3preciocom1.Text;
        String preciocom2 = li3preciocom2.Text;
        String facingdesa = li3facingdesa.Text;
        String facingcom1 = li3facingcom1.Text;
        String facingcom2 = li3facingcom2.Text;
        String dispdesa = ddi3disdesa.SelectedValue.ToString();
        String dispcom1 = ddi3discom1.SelectedValue.ToString();
        String dispcom2 = ddi3discom2.SelectedValue.ToString();
        String promdesa = ddi3promdesa.SelectedValue.ToString();
        String promcom1 = ddi3promcom1.SelectedValue.ToString();
        String promcom2 = ddi3promcom2.SelectedValue.ToString();
        String exdesa = ddi3exdesa.SelectedValue.ToString();
        String excom1 = ddi3excom1.SelectedValue.ToString();
        String excom2 = ddi3excom2.SelectedValue.ToString();
        Boolean validador = false;

        cnn.Open();

        SQL = " INSERT INTO [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
               " VALUES ('" + idencuestador + "', '" + idpdv + "', '" + fecha + "', '" + hora + "', '" + iditem + "','" + preciodesa + "', " +
               " '" + preciocom1 + "','" + preciocom2 + "','" + facingdesa + "','" + facingcom1 + "','" + facingcom2 + "','" + dispdesa + "','" + dispcom1 + "', " +
               " '" + dispcom2 + "','" + promdesa + "','" + promcom1 + "','" + promcom2 + "','" + exdesa + "','" + excom1 + "','" + excom2 + "');  " +
               " SELECT * FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] WHERE idrespuesta = SCOPE_IDENTITY(); ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            validador = true;
        }
        else
        {
            validador = false;
        }

        reader.Close();
        cnn.Close();

        return validador;
    }

    //Se genera el Ingreso de la cuarta encuesta si corresponde...
    protected Boolean Insertaitem4()
    {
        String idencuestador = Session["idusuario"].ToString();
        String idpdv = Session["idpdv"].ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String hora = DateTime.Now.ToString("HHmmss");
        String iditem = li4versus.Text;
        String preciodesa = li4preciodesa.Text;
        String preciocom1 = li4preciocom1.Text;
        String preciocom2 = li4preciocom2.Text;
        String facingdesa = li4facingdesa.Text;
        String facingcom1 = li4facingcom1.Text;
        String facingcom2 = li4facingcom2.Text;
        String dispdesa = ddi4disdesa.SelectedValue.ToString();
        String dispcom1 = ddi4discom1.SelectedValue.ToString();
        String dispcom2 = ddi4discom2.SelectedValue.ToString();
        String promdesa = ddi4promdesa.SelectedValue.ToString();
        String promcom1 = ddi4promcom1.SelectedValue.ToString();
        String promcom2 = ddi4promcom2.SelectedValue.ToString();
        String exdesa = ddi4exdesa.SelectedValue.ToString();
        String excom1 = ddi4excom1.SelectedValue.ToString();
        String excom2 = ddi4excom2.SelectedValue.ToString();
        Boolean validador = false;

        cnn.Open();

        SQL = " INSERT INTO [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
               " VALUES ('" + idencuestador + "', '" + idpdv + "', '" + fecha + "', '" + hora + "', '" + iditem + "','" + preciodesa + "', " +
               " '" + preciocom1 + "','" + preciocom2 + "','" + facingdesa + "','" + facingcom1 + "','" + facingcom2 + "','" + dispdesa + "','" + dispcom1 + "', " +
               " '" + dispcom2 + "','" + promdesa + "','" + promcom1 + "','" + promcom2 + "','" + exdesa + "','" + excom1 + "','" + excom2 + "');  " +
               " SELECT * FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] WHERE idrespuesta = SCOPE_IDENTITY(); ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            validador = true;
        }
        else
        {
            validador = false;
        }

        reader.Close();
        cnn.Close();

        return validador;
    }

    //Se genera el Ingreso de la quinta encuesta si corresponde...
    protected Boolean Insertaitem5()
    {
        String idencuestador = Session["idusuario"].ToString();
        String idpdv = Session["idpdv"].ToString();
        String fecha = DateTime.Now.ToString("yyyyMMdd");
        String hora = DateTime.Now.ToString("HHmmss");
        String iditem = li5versus.Text;
        String preciodesa = li5preciodesa.Text;
        String preciocom1 = li5preciocom1.Text;
        String preciocom2 = li5preciocom2.Text;
        String facingdesa = li5facingdesa.Text;
        String facingcom1 = li5facingcom1.Text;
        String facingcom2 = li5facingcom2.Text;
        String dispdesa = ddi5disdesa.SelectedValue.ToString();
        String dispcom1 = ddi5discom1.SelectedValue.ToString();
        String dispcom2 = ddi5discom2.SelectedValue.ToString();
        String promdesa = ddi5promdesa.SelectedValue.ToString();
        String promcom1 = ddi5promcom1.SelectedValue.ToString();
        String promcom2 = ddi5promcom2.SelectedValue.ToString();
        String exdesa = ddi5exdesa.SelectedValue.ToString();
        String excom1 = ddi5excom1.SelectedValue.ToString();
        String excom2 = ddi5excom2.SelectedValue.ToString();
        Boolean validador = false;

        cnn.Open();

        SQL = " INSERT INTO [DataWarehouse].[dbo].[ENC_RESPUESTA] " +
               " VALUES ('" + idencuestador + "', '" + idpdv + "', '" + fecha + "', '" + hora + "', '" + iditem + "','" + preciodesa + "', " +
               " '" + preciocom1 + "','" + preciocom2 + "','" + facingdesa + "','" + facingcom1 + "','" + facingcom2 + "','" + dispdesa + "','" + dispcom1 + "', " +
               " '" + dispcom2 + "','" + promdesa + "','" + promcom1 + "','" + promcom2 + "','" + exdesa + "','" + excom1 + "','" + excom2 + "');  " +
               " SELECT * FROM [DataWarehouse].[dbo].[ENC_RESPUESTA] WHERE idrespuesta = SCOPE_IDENTITY(); ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        if (reader.Read())
        {
            validador = true;
        }
        else
        {
            validador = false;
        }

        reader.Close();
        cnn.Close();

        return validador;
    }

    //Validación segun la info en la tabla de item (Precio)
    protected Boolean Validar()
    {
        Boolean val=true;
        String iditem1 = li1versus.Text;
        String preciodesa1 = li1preciodesa.Text;
        String iditem2 = li2versus.Text;
        String preciodesa2 = li2preciodesa.Text;
        String iditem3 = li3versus.Text;
        String preciodesa3 = li3preciodesa.Text;
        String iditem4 = li4versus.Text;
        String preciodesa4 = li4preciodesa.Text;
        String iditem5 = li5versus.Text;
        String preciodesa5 = li5preciodesa.Text;

        cnn.Open();

        SQL = " SELECT [iditem],[pdesadesde],[pdesahasta],[nombredesa] " +
              " FROM [DataWarehouse].[dbo].[ENC_ITEM] " +
              " Where pdesadesde!=0 and pdesahasta!=0 ";

        cmd = new SqlCommand(SQL, cnn);
        reader = cmd.ExecuteReader();

        while (reader.Read())
        {
            //Corrobora el Item 1
            if (iditem1.Equals(reader.GetInt32(0).ToString()) && Convert.ToInt32(preciodesa1)!=0 && (reader.GetInt32(1) > Convert.ToInt32(preciodesa1) || reader.GetInt32(2) < Convert.ToInt32(preciodesa1)))
            {
                val = false;
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Error", "alert('Error al ingresar la encuesta, el precio de " + reader.GetString(3) + " debe estar entre " + reader.GetInt32(1).ToString() + " y " + reader.GetInt32(2).ToString() + "');", true);
            }

            //Corrobora el Item 2
            if (iditem2.Equals(reader.GetInt32(0).ToString()) && Convert.ToInt32(preciodesa2) != 0 && (reader.GetInt32(1) > Convert.ToInt32(preciodesa2) || reader.GetInt32(2) < Convert.ToInt32(preciodesa2)))
            {
                val = false;
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Error", "alert('Error al ingresar la encuesta, el precio de " + reader.GetString(3) + " debe estar entre " + reader.GetInt32(1).ToString() + " y " + reader.GetInt32(2).ToString() + "');", true);
            }

            //Corrobora el Item 3
            if (iditem3.Equals(reader.GetInt32(0).ToString()) && Convert.ToInt32(preciodesa3) != 0 && (reader.GetInt32(1) > Convert.ToInt32(preciodesa3) || reader.GetInt32(2) < Convert.ToInt32(preciodesa3)))
            {
                val = false;
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Error", "alert('Error al ingresar la encuesta, el precio de " + reader.GetString(3) + " debe estar entre " + reader.GetInt32(1).ToString() + " y " + reader.GetInt32(2).ToString() + "');", true);
            }

            //Corrobora el Item 4
            if (iditem4.Equals(reader.GetInt32(0).ToString()) && Convert.ToInt32(preciodesa4) != 0 && (reader.GetInt32(1) > Convert.ToInt32(preciodesa4) || reader.GetInt32(2) < Convert.ToInt32(preciodesa4)))
            {
                val = false;
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Error", "alert('Error al ingresar la encuesta, el precio de " + reader.GetString(3) + " debe estar entre " + reader.GetInt32(1).ToString() + " y " + reader.GetInt32(2).ToString() + "');", true);
            }

            //Corrobora el Item 5
            if (iditem5.Equals(reader.GetInt32(0).ToString()) && Convert.ToInt32(preciodesa5) != 0 && (reader.GetInt32(1) > Convert.ToInt32(preciodesa5) || reader.GetInt32(2) < Convert.ToInt32(preciodesa5)))
            {
                val = false;
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Error", "alert('Error al ingresar la encuesta, el precio de " + reader.GetString(3) + " debe estar entre " + reader.GetInt32(1).ToString() + " y " + reader.GetInt32(2).ToString() + "');", true);
            }
        }

        reader.Close();
        cnn.Close();

        return val;
    }

    //Refresco para nueva encuesta
    protected void Npdv_Click(object sender, EventArgs e)
    {
        Session["idpdv"] = "";
        Response.Redirect(Request.RawUrl);
    }
}