using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
//using System.Data.SqlClient.SqlCommand;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Inf_Semanal
{
    public partial class Form1 : Form
    {
        
        public  string  url_archivo;
        public int id_mes;
        /*
        public string string_conexion = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
          @"Data source= C:\Users\MC01\Documents\Visual Studio 2013\Projects\Inf_Semanal\Inf_Semanal\bin\Debug\" +
          @"Basedatosinfsemanal.mdb";*/

        public string string_conexion = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
          @"Data source= \\Mc01\mc01\basededatos\" +
          @"Basedatosinfsemanal.mdb";
        public Form1()
        {
            InitializeComponent();
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
             
            }


       
        private void button2_Click(object sender, EventArgs e)
        {
          
            if (txt_user.Text != "" || txt_password.Text != "")
            {
                
                try
                {
                    OleDbConnection conexion = new OleDbConnection(string_conexion);
                    OleDbCommand selec_usu = new OleDbCommand(string.Format("SELECT acceso.id, acceso.nombres, acceso.apellidos, acceso.id_tipo, acceso.id_tipo_usuario, tipo_usuario.tipo FROM acceso, tipo_usuario  WHERE acceso.usuarios ='{0}' AND acceso.password='{1}' AND  acceso.id_tipo_usuario=tipo_usuario.id", txt_user.Text, txt_password.Text), conexion);

                    conexion.Open();

                    OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();
           
                    if (lectorRegistros.HasRows)
                    {
                       
                        while (lectorRegistros.Read())
                        {
                        
                            string[] datos_user= new string[6];
                           
                            datos_user[0] = lectorRegistros["nombres"].ToString();
                            datos_user[1] = lectorRegistros["apellidos"].ToString();
                            datos_user[2] = lectorRegistros["id"].ToString();
                            datos_user[3] = lectorRegistros["id_tipo"].ToString();
                            datos_user[4] = lectorRegistros["id_tipo_usuario"].ToString();
                            datos_user[5] = lectorRegistros["tipo"].ToString();
                            if(datos_user[4]=="2"){
                                Realizar_formato nuevo_form = new Realizar_formato(datos_user, id_mes);
                                this.Hide();
                                txt_user.Text = "";
                                txt_password.Text = "";
                                nuevo_form.ShowDialog();
                                this.Show();
                                return;
                            }
                            Administrador administra = new Administrador(datos_user);
                            this.Hide();
                            txt_user.Text = "";
                            txt_password.Text = "";
                            administra.ShowDialog();
                            this.Show();


                        }
                    }
                    else
                    {
                        MessageBox.Show("Usuario No Registrado");
                    }




                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else {
                MessageBox.Show("Digite su Usuario o Contraseña");
            }
            

            }

        private void Form1_EnabledChanged(object sender, EventArgs e)
        {
           
        }

        private void txt_user_KeyPress(object sender, KeyPressEventArgs e)
       {
            if (!(char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten letras", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

      
        private void txt_password_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten letras", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_password_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_user_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Formcsv formula = new Formcsv();
            formula.Show();
        }
    }
}
