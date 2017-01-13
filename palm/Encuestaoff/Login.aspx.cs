using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;

public partial class Login : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        SqlConnection cnn1 = new SqlConnection("Data Source=SQLSERVER;Initial Catalog=DesaERP;Persist Security Info=True;User ID=sa;password=desakey");
        cnn1.Open();

        String sql1 = " SELECT [idencuestador],[nombre] FROM [DataWarehouse].[dbo].[ENC_ENCUESTADORES] ";
        SqlCommand cmd1 = new SqlCommand(sql1, cnn1);
        SqlDataReader reader1 = cmd1.ExecuteReader();

        while (reader1.Read())
        {
            ddusuarios.Items.Add(new ListItem(reader1.GetInt32(0).ToString() + ".- " + reader1.GetString(1),reader1.GetInt32(0).ToString()));
        }

        reader1.Close();
        cnn1.Close();
    }

    protected void Acceder_Click(object sender, EventArgs e)
    {
        String Usuario = ddusuarios.SelectedValue.ToString();
        String Pass = LPass.Text;

        SqlConnection cnn1 = new SqlConnection("Data Source=SQLSERVER;Initial Catalog=DesaERP;Persist Security Info=True;User ID=sa;password=desakey");
        cnn1.Open();

        String sql1 = " SELECT [idencuestador],[nombre] FROM [DataWarehouse].[dbo].[ENC_ENCUESTADORES] " +
                      " Where UPPER(idencuestador)='" + Usuario.ToUpper() + "' and UPPER(contraseña)='" + Pass.ToUpper() + "'";

        SqlCommand cmd1 = new SqlCommand(sql1, cnn1);
        SqlDataReader reader1 = cmd1.ExecuteReader();

        if (reader1.Read())
        {
            Session["usuario"] = reader1.GetString(1);
            Session["idusuario"] = reader1.GetInt32(0);
            Response.Redirect("http://pda.desa.cl/palm/encuestaoff/Encuesta.aspx");
        }
        else
        {
            string script = "<script language='javascript'>alert('Contraseña erronea!!!');</script>";
            this.ClientScript.RegisterStartupScript(this.GetType(), Guid.NewGuid().ToString(), script);
        }

        reader1.Close();
        cnn1.Close();
    }
}