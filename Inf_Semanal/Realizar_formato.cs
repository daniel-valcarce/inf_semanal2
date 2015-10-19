using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace Inf_Semanal
{
    public partial class Realizar_formato : Form
    {

        public string url_archivo;
        public string mes_selec;
        public int id_mes;
        public OleDbDataReader informe_extraido;
        public OleDbDataReader metas;
        public OleDbDataReader porciento_cumplimiento;
        public OleDbDataReader comentario;
        public Usuario usuario1 = new Usuario();
        public Microsoft.Office.Interop.Excel.Workbook documentoexcel0;

        public Realizar_formato(string[] datos_user, int mes_selec)
        {
            //ACa en la linea de abajo llamo el metodo para traer los meses al combobox
            System.Data.DataTable meses = CRUD.selec_meses();
            //aca asigno el id de mes seleccionado q me llega en la carga
            this.id_mes= mes_selec;
            // creo una instancia de la clase usuario con los datos del array q llegan
            usuario1.nombres = datos_user[0];
            usuario1.apellidos = datos_user[1];
            usuario1.id_user = Convert.ToInt32(datos_user[2]);
            usuario1.tipo_asesor = Convert.ToInt32(datos_user[3]);
            usuario1.tipo_usuario = Convert.ToInt32(datos_user[4]);
            usuario1.cadena_tipo_usuario = datos_user[5];
            //Traigo las metas
            metas = CRUD.obtener_Metas(usuario1.tipo_asesor);
          
            //inicializo componentes
            InitializeComponent();
            //Coloco las metas en los labels
            colocar_metas_en_labels(usuario1.tipo_asesor);
            //inhabilito lo q tenga q desabilitar
            btn_visualizar.Enabled = false;
            panel1.Enabled = false;
            panel2.Enabled=false;
            panel3.Enabled=false;
            btn_cancelar.Enabled = false;
            btn_listo.Enabled = false;
            btn_realizar_form.Enabled = false;
            btn_act_form.Enabled = false;
            btn_act_infor_cons.Enabled = false;
                //decido que se va a mostrar

            if(usuario1.tipo_asesor==3){
                panel3.Hide();
                }

            //asigno valores en combobox
            cmb_mes.DisplayMember = "MES";
            cmb_mes.ValueMember = "ID";
            cmb_mes.DataSource = meses;
            //mostramos datos de usuario en los labels
            lbl_nomb_asesor.Text = "" + usuario1.nombres + " " + usuario1.apellidos + "";
            lbl_tipo_usuario.Text = usuario1.cadena_tipo_usuario;
            //
            if(this.id_mes==0){
                DateTime fechaActual = DateTime.Now;
                this.id_mes = fechaActual.Month;    
            }

            cmb_mes.SelectedIndex = id_mes-1;
            //termino de asignar valores en combobox
            System.Data.DataTable semana = CRUD.selec_semanas(Convert.ToInt32(cmb_mes.SelectedValue));
            cmb_semana.DisplayMember = "SEMANA";
            cmb_semana.ValueMember = "ID";
            cmb_semana.DataSource = semana;
            cmb_semana.SelectedItem = null;
        }

        public void colocar_metas_en_labels(int tipo_asesor) {
          
            while(metas.Read()){
               
   lbl_met_cemen_boq.Text = (Convert.ToInt32(metas["cemento_boquilla"].ToString())).ToString();
   lbl_met_cenefa.Text = (Convert.ToInt32(metas["cenefa"].ToString())).ToString();
   lbl_met_cera.Text = (Convert.ToInt32(metas["ceramica"].ToString())).ToString();

   
   lbl_met_grife.Text = (Convert.ToInt32(metas["griferia"].ToString()) ).ToString();
   lbl_met_lavapla.Text = (Convert.ToInt32(metas["lavaplatos"].ToString())).ToString();
   lbl_met_pegos.Text = (Convert.ToInt32(metas["pegos"].ToString())).ToString();
   
   lbl_met_pint.Text = (Convert.ToInt32(metas["pintura"].ToString())).ToString();
   lbl_met_porcela.Text = (Convert.ToInt32(metas["porcelanato"].ToString())).ToString();
   
   lbl_met_sanita.Text = (Convert.ToInt32(metas["sanitarios"].ToString())).ToString();
  
   lbl_met_lijas.Text = (Convert.ToInt32(metas["lijas"].ToString())).ToString();
   lbl_met_brochass.Text = (Convert.ToInt32(metas["brochas"].ToString())).ToString();

   lbl_met_tube.Text = (Convert.ToInt32(metas["tuberia"].ToString())).ToString();
   
   
   lbl_met_incrust.Text = (Convert.ToInt32(metas["incrustaciones"].ToString())).ToString();

                if(tipo_asesor!=3){
                    lbl_met_tapetes.Text = (Convert.ToInt32(metas["tapetes"].ToString())).ToString();
                    lbl_met_gabin.Text = (Convert.ToInt32(metas["gabinete_espejos"].ToString())).ToString();
                    lbl_met_gres.Text = (Convert.ToInt32(metas["gres"].ToString())).ToString();
                    lbl_met_adhesivos.Text = (Convert.ToInt32(metas["adhesivos"].ToString())).ToString();
                    lbl_met_perfiles.Text = (Convert.ToInt32(metas["perfiles"].ToString())).ToString();
                    lbl_met_electri.Text = (Convert.ToInt32(metas["electrico"].ToString())).ToString();
                    lbl_met_cocinas.Text = (Convert.ToInt32(metas["cocinas_integrales"].ToString())).ToString();
                    lbl_met_produQui.Text = (Convert.ToInt32(metas["produc_quimi"].ToString())).ToString();
                }
               
                   int total = ((Convert.ToInt32(metas["cemento_boquilla"].ToString()) +
                Convert.ToInt32(metas["cenefa"].ToString())+
                Convert.ToInt32(metas["ceramica"].ToString())+
                Convert.ToInt32(metas["griferia"].ToString()) +
                Convert.ToInt32(metas["lavaplatos"].ToString()) +
                Convert.ToInt32(metas["pegos"].ToString()) +
                Convert.ToInt32(metas["pintura"].ToString()) +
                Convert.ToInt32(metas["porcelanato"].ToString()) +
                Convert.ToInt32(metas["sanitarios"].ToString()) +
                Convert.ToInt32(metas["lijas"].ToString()) +
                Convert.ToInt32(metas["brochas"].ToString()) +
                Convert.ToInt32(metas["tuberia"].ToString()) +
                Convert.ToInt32(metas["incrustaciones"].ToString())));
                   if (tipo_asesor != 3) {
                      int  mas = Convert.ToInt32(metas["tapetes"].ToString()) +
                    Convert.ToInt32(metas["gabinete_espejos"].ToString()) +
                    Convert.ToInt32(metas["gres"].ToString()) +
                    Convert.ToInt32(metas["adhesivos"].ToString()) +
                    Convert.ToInt32(metas["perfiles"].ToString()) +
                    Convert.ToInt32(metas["electrico"].ToString()) +
                    Convert.ToInt32(metas["cocinas_integrales"].ToString()) +
                    Convert.ToInt32(metas["produc_quimi"].ToString());
                      total = (total + ((mas)));
                   }
                   lbl_total.Text =total.ToString();
                   lbl_total_semana.Text = (total / 4).ToString();
            }

        }

        private void btn_conf_mes_Click(object sender, EventArgs e)
        {
            
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            btn_listo.Enabled = true;
            btn_cancelar.Enabled = true;
            btn_act_infor_cons.Enabled = false;
            panel1.Enabled = true;
            panel2.Enabled = true;
            panel3.Enabled = true;
            txt_ventas_semanal.Enabled = true;
            porciento_cumplimiento = CRUD.dias_habiles(Convert.ToInt32(cmb_mes.SelectedValue),Convert.ToInt32(cmb_semana.SelectedValue));


        }

        private void Realizar_formato_FormClosed(object sender, FormClosedEventArgs e)
        {
            
           
         
           
        }


        

        private void txt_cemen_boqui_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_cenefa_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_ceramica_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_griferia_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_lavaplatos_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_pegos_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_pintura_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_porcela_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_sanitario_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_brochas_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_lijas_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_tuberia_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_incrus_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_tapetes_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_gabine_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_gres_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

    

        private void txt_adhesivos_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_perfiles_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_electrico_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_cocinas_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_prod_quim_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void cmb_semana_SelectionChangeCommitted(object sender, EventArgs e)
        {
            panel1.Enabled = false; panel2.Enabled = false; panel3.Enabled = false;
            btn_cancelar.Enabled = false; btn_listo.Enabled = false;

            limpiartexto();
            comentario = CRUD.dias_habiles(Convert.ToInt32(cmb_mes.SelectedValue), Convert.ToInt32(cmb_semana.SelectedValue));
            while(comentario.Read()){
                lbl_coment_semana.Text = comentario["comentario"].ToString();
                lbl_porc_cumpli_ideal.Text = (Convert.ToDecimal(comentario["dias_habiles_actual"].ToString()) / Convert.ToDecimal(comentario["dias_habiles_mes"].ToString())).ToString("P");
            }
            informe_extraido = CRUD.Determina_Oper(usuario1.id_user, Convert.ToInt32(cmb_mes.SelectedValue), Convert.ToInt32(cmb_semana.SelectedValue), usuario1.tipo_asesor);
            if (informe_extraido == null)
            {
                btn_realizar_form.Enabled = true;
                btn_act_form.Enabled = false;
                btn_visualizar.Enabled = false;

            } if (informe_extraido != null)
            {
                btn_act_form.Enabled = true;
                btn_visualizar.Enabled = true;
                btn_realizar_form.Enabled = false;
            }
           
        }
        public  void limpiartexto() {
            txt_cemen_boqui.Text = "";
            txt_cenefa.Text = "";
            txt_ceramica.Text = "";
            txt_griferia.Text = "";
            txt_lavaplatos.Text = "";
            txt_pegos.Text = "";
            txt_pintura.Text = "";
            txt_porcela.Text = "";
            txt_sanitario.Text = "";
            txt_lijas.Text = "";
            txt_brochas.Text = "";
            txt_tuberia.Text = "";
            txt_incrus.Text = "";
            txt_tapetes.Text = "";
            txt_gabine.Text = "";
            txt_gres.Text = "";
            txt_adhesivos.Text = "";
            txt_perfiles.Text = "";
            txt_electrico.Text = "";
            txt_cocinas.Text = "";
            txt_prod_quim.Text = "";
            txt_ventas_semanal.Text = "";
            lbl_ventas_totales.Text = "";
            lbl_porc_cumpl_semana.Text = "";
            lbl_porc_cumpli_ideal.Text = "";
            lbl_porc_cumpli_real.Text = "";
            lbl_coment_semana.Text = "";
            lbl_porc_cenefa.Text = "";
            lbl_porc_cem_boqui.Text = "";
            lbl_porc_cera.Text = "";
            lbl_porc_cocinas.Text = "";
            lbl_porc_incrusta.Text = "";
            lbl_porc_griferia.Text = "";
            lbl_porc_gres.Text = "";
            lbl_porc_gabine.Text = "";
            lbl_porc_lavapla.Text = "";
            lbl_porc_porcela.Text = "";
            lbl_porc_adhesivo.Text = "";
            lbl_porc_brochas.Text="";
            lbl_porc_pintura.Text = "";
            lbl_porc_perfil.Text = "";
            lbl_porc_sanita.Text = "";
            lbl_porc_tapetes.Text = "";
            lbl_porc_pego.Text = "";
            lbl_porc_lijas.Text = "";
            lbl_porc_tuberia.Text = "";
            lbl_porc_prodQuimi.Text="";
            lbl_porc_electrico.Text = "";
        
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            limpiartexto();
            informe_extraido = null;
            btn_listo.Enabled = false;
            btn_cancelar.Enabled = false;
            cmb_semana.SelectedItem = false;
            btn_visualizar.Enabled = false;
            btn_realizar_form.Enabled = false;
            btn_act_form.Enabled = false;
            btn_act_infor_cons.Enabled = false;
            txt_ventas_semanal.Enabled = false;
        }


        public void actuaizar_visualizar() {

            btn_listo.Enabled = false;
            btn_cancelar.Enabled = true;
            btn_act_infor_cons.Enabled = true;
            panel1.Enabled = true;
            panel2.Enabled = true;
            panel3.Enabled = true;
            txt_ventas_semanal.Enabled = true;
            while (informe_extraido.Read())
            {
                txt_cemen_boqui.Text = informe_extraido["CEMENTO_BOQUILLA"].ToString();
                txt_cenefa.Text = informe_extraido["CENEFA"].ToString();
                txt_ceramica.Text = informe_extraido["CERAMICA"].ToString();


                txt_griferia.Text = informe_extraido["GRIFERIA"].ToString();
                txt_lavaplatos.Text = informe_extraido["LAVAPLATOS"].ToString();
                txt_pegos.Text = informe_extraido["PEGOS"].ToString();

                txt_pintura.Text = informe_extraido["PINTURA"].ToString();
                txt_porcela.Text = informe_extraido["PORCELANATO"].ToString();

                txt_sanitario.Text = informe_extraido["SANITARIOS"].ToString();

                txt_lijas.Text = informe_extraido["LIJAS"].ToString();
                txt_brochas.Text = informe_extraido["BROCHAS"].ToString();

                txt_tuberia.Text = informe_extraido["TUBERIA"].ToString();

                txt_incrus.Text = informe_extraido["INCRUSTACIONES"].ToString();

                txt_ventas_semanal.Text = informe_extraido["VENTAS_SEMANAL"].ToString();

                decimal total_ventas =Convert.ToDecimal(txt_cemen_boqui.Text)+Convert.ToDecimal(txt_cenefa.Text)+Convert.ToDecimal( txt_ceramica.Text)+Convert.ToDecimal(txt_griferia.Text)+Convert.ToDecimal(txt_lavaplatos.Text)+Convert.ToDecimal(txt_pegos.Text)+Convert.ToDecimal(txt_pintura.Text)+Convert.ToDecimal(txt_porcela.Text)+Convert.ToDecimal(txt_sanitario.Text)+Convert.ToDecimal(txt_lijas.Text)+Convert.ToDecimal(txt_brochas.Text)+Convert.ToDecimal(txt_tuberia.Text)+Convert.ToDecimal(txt_incrus.Text);
                if (usuario1.tipo_asesor == 1 || usuario1.tipo_asesor == 2)
                {

                    txt_tapetes.Text = informe_extraido["TAPETES"].ToString();
                    txt_gabine.Text = informe_extraido["GABINETES_ESPEJOS"].ToString();
                    txt_gres.Text = informe_extraido["GRES"].ToString();
                    txt_adhesivos.Text = informe_extraido["ADHESIVOS"].ToString();
                    txt_perfiles.Text = informe_extraido["PERFILES"].ToString();
                    txt_electrico.Text = informe_extraido["ELECTRICO"].ToString();
                    txt_cocinas.Text = informe_extraido["COCINAS_INTREGALES"].ToString();
                    txt_prod_quim.Text = informe_extraido["PRODUCTOS_QUIMICOS"].ToString();

                    decimal mas_total = Convert.ToDecimal(txt_tapetes.Text) + Convert.ToDecimal(txt_gabine.Text) + Convert.ToDecimal(txt_gres.Text) + Convert.ToDecimal(txt_adhesivos.Text) + Convert.ToDecimal(txt_perfiles.Text) + Convert.ToDecimal(txt_electrico.Text) + Convert.ToDecimal(txt_cocinas.Text) + Convert.ToDecimal(txt_prod_quim.Text);
                   total_ventas = total_ventas + mas_total;
                }

                lbl_ventas_totales.Text = total_ventas.ToString();
                lbl_porc_cumpli_real.Text = (total_ventas / Convert.ToDecimal(lbl_total.Text)).ToString("P");
                lbl_porc_cumpl_semana.Text = (Convert.ToDecimal(txt_ventas_semanal.Text) / Convert.ToDecimal(lbl_total_semana.Text)).ToString("P");
            }
        
        }



        private void btn_act_form_Click(object sender, EventArgs e)
        {
            limpiartexto();
            btn_realizar_form.Enabled = false;
            btn_visualizar.Enabled = false;
            actuaizar_visualizar();
            porcientos();

            
        }


        private void cmb_semana_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public  double[] tomar_valores()
        {
            double[] valores_informe = new double[22];
            if (txt_cemen_boqui.Text == "" ||txt_ventas_semanal.Text==""|| txt_cenefa.Text == "" || txt_ceramica.Text == "" || txt_griferia.Text == "" || txt_lavaplatos.Text == "" || txt_pegos.Text == "" || txt_pintura.Text == "" || txt_porcela.Text == "" || txt_sanitario.Text == "" || txt_lijas.Text == "" || txt_brochas.Text == "" || txt_tuberia.Text == "" || txt_incrus.Text=="")
            
            {
                
                return valores_informe = null;
            }

           
            valores_informe[0] = double.Parse(txt_cemen_boqui.Text);
            valores_informe[1] = double.Parse(txt_cenefa.Text);
            valores_informe[2] = double.Parse(txt_ceramica.Text);   
            valores_informe[5] = double.Parse(txt_griferia.Text);
            valores_informe[6] = double.Parse(txt_lavaplatos.Text);
            valores_informe[7] = double.Parse(txt_pegos.Text);
            valores_informe[9] = double.Parse(txt_pintura.Text);
            valores_informe[10] = double.Parse(txt_porcela.Text);  
            valores_informe[12] = double.Parse(txt_sanitario.Text);
            valores_informe[14] = double.Parse(txt_lijas.Text);
            valores_informe[15] = double.Parse(txt_brochas.Text);
            valores_informe[17] = double.Parse(txt_tuberia.Text);
            valores_informe[20] = double.Parse(txt_incrus.Text);
            valores_informe[21] = double.Parse(txt_ventas_semanal.Text);
            if(usuario1.tipo_asesor!=3){

                if (txt_tapetes.Text == "" || txt_gabine.Text == "" || txt_gres.Text == "" || txt_adhesivos.Text == "" || txt_perfiles.Text == "" || txt_electrico.Text == "" || txt_cocinas.Text == "" || txt_prod_quim.Text=="")
                {
                    return valores_informe = null;
                }
                valores_informe[13] = double.Parse(txt_tapetes.Text);
                valores_informe[3] = double.Parse(txt_gabine.Text);
                valores_informe[4] = double.Parse(txt_gres.Text);
                valores_informe[16] = double.Parse(txt_adhesivos.Text);
                valores_informe[8] = double.Parse(txt_perfiles.Text);
                valores_informe[18] = double.Parse(txt_electrico.Text);
                valores_informe[19] = double.Parse(txt_cocinas.Text);
                valores_informe[11] = double.Parse(txt_prod_quim.Text);


            }
            return valores_informe;
        }

    






        private void btn_listo_Click(object sender, EventArgs e)
        {
           
            double[] valores_informe = tomar_valores();

             if(valores_informe==null){
          MessageBox.Show("Ingrese el valor de ventas en las casillas, o coloque 0");
          return;
             }

            int respuesta_insert = CRUD.insert_infor_semanal(usuario1.id_user, Convert.ToInt32(cmb_mes.SelectedValue), Convert.ToInt32(cmb_semana.SelectedValue), usuario1.tipo_asesor, valores_informe);
            if (respuesta_insert == 1)
            {
                MessageBox.Show("Formato Realizado Correctamente");
                string hoja = cmb_semana.GetItemText(cmb_semana.SelectedItem);
              
                cmb_semana.SelectedItem = null;
                

            }
            else {
                MessageBox.Show("Error al Realizar el Formato, Verifica Los Datos");
            }
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btn_act_infor_cons_Click(object sender, EventArgs e)
        {
           
            double[] valores_informe = tomar_valores();

            int respuesta_insert = CRUD.update_informe(usuario1.id_user, Convert.ToInt32(cmb_mes.SelectedValue), Convert.ToInt32(cmb_semana.SelectedValue), usuario1.tipo_asesor, valores_informe);
            if (respuesta_insert == 1)
            {
                MessageBox.Show("Formato Realizado Correctamente");
                string hoja = cmb_semana.GetItemText(cmb_semana.SelectedItem);

                cmb_semana.SelectedItem = null;
                
            }
            else
            {
                MessageBox.Show("Error al Realizar el Formato, Verifica Los Datos");
            }
        }

        private void cmb_mes_SelectionChangeCommitted(object sender, EventArgs e)
        {
            limpiartexto();
            informe_extraido = null;
            btn_listo.Enabled = false;
            btn_cancelar.Enabled = false;
            cmb_semana.SelectedItem = false;
            btn_visualizar.Enabled = false;
            btn_realizar_form.Enabled = false;
            btn_act_form.Enabled = false;
            btn_act_infor_cons.Enabled = false;
            txt_ventas_semanal.Enabled = false;


            lbl_coment_semana.Text = "";
            System.Data.DataTable semana = CRUD.selec_semanas(Convert.ToInt32(cmb_mes.SelectedValue));
            cmb_semana.DisplayMember = "SEMANA";
            cmb_semana.ValueMember = "ID";
            cmb_semana.DataSource = semana;
            cmb_semana.SelectedItem = null;
        }

        private void cmb_mes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cmb_semana.SelectedItem!=null){
                panel1.Enabled = false; panel2.Enabled = false; panel3.Enabled = false;
                btn_cancelar.Enabled = false; btn_listo.Enabled = false;
                limpiartexto();
                informe_extraido = CRUD.Determina_Oper(usuario1.id_user, Convert.ToInt32(cmb_mes.SelectedValue), Convert.ToInt32(cmb_semana.SelectedValue), usuario1.tipo_asesor);
                if (informe_extraido == null)
                {
                    btn_realizar_form.Enabled = true;
                    btn_act_form.Enabled = false;
                    btn_visualizar.Enabled = false;

                } if (informe_extraido != null)
                {
                    btn_act_form.Enabled = true;
                    btn_visualizar.Enabled = true;
                    btn_realizar_form.Enabled = false;
                }
            
            }

            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Enero")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\ENERO\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Febrero")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\FEBRERO\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Marzo")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\MARZO\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Abril")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\ABRIL\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Mayo")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\MAYO\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Junio")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\JUNIO\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Julio")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\JULIO\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Agosto")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\AGOSTO\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Septiembre")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\SEPTIEMBRE\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Octubre")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\OCTUBRE\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Noviembre")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\NOVIEMBRE\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
            if (cmb_mes.GetItemText(cmb_mes.SelectedItem) == "Diciembre")
            {
                url_archivo = "\\subgerente\\sistema de gestion de calidad\\PROCESO COMERCIAL\\PROCESO COMERCIAL\\REGISTROS 2015\\SEGUIMIENTO\\DICIEMBRE\\INFORME GENERAL DE VENTAS SEMANAL.xlsx";
            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            actuaizar_visualizar();
            btn_act_infor_cons.Enabled = false;
            btn_act_form.Enabled = false;
            btn_listo.Enabled = false;
            btn_cancelar.Enabled = true;
            porcientos();
           
        }


        public void porcientos()
        {
            lbl_porc_cem_boqui.Text = (Convert.ToDecimal(txt_cemen_boqui.Text) / Convert.ToDecimal(lbl_met_cemen_boq.Text)).ToString("P");
            lbl_porc_cenefa.Text = (Convert.ToDecimal(txt_cenefa.Text) / Convert.ToDecimal(lbl_met_cenefa.Text)).ToString("P");
            lbl_porc_cera.Text = (Convert.ToDecimal(txt_ceramica.Text) / Convert.ToDecimal(lbl_met_cera.Text)).ToString("P");
            lbl_porc_griferia.Text = (Convert.ToDecimal(txt_griferia.Text) / Convert.ToDecimal(lbl_met_grife.Text)).ToString("P");
            lbl_porc_lavapla.Text = (Convert.ToDecimal(txt_lavaplatos.Text) / Convert.ToDecimal(lbl_met_lavapla.Text)).ToString("P");
            lbl_porc_pego.Text = (Convert.ToDecimal(txt_pegos.Text) / Convert.ToDecimal(lbl_met_pegos.Text)).ToString("P");
            lbl_porc_pintura.Text = (Convert.ToDecimal(txt_pintura.Text) / Convert.ToDecimal(lbl_met_pint.Text)).ToString("P");
            lbl_porc_porcela.Text = (Convert.ToDecimal(txt_porcela.Text) / Convert.ToDecimal(lbl_met_porcela.Text)).ToString("P");
            lbl_porc_sanita.Text = (Convert.ToDecimal(txt_sanitario.Text) / Convert.ToDecimal(lbl_met_sanita.Text)).ToString("P");
            lbl_porc_brochas.Text = (Convert.ToDecimal(txt_brochas.Text) / Convert.ToDecimal(lbl_met_brochass.Text)).ToString("P");
            lbl_porc_lijas.Text = (Convert.ToDecimal(txt_lijas.Text) / Convert.ToDecimal(lbl_met_lijas.Text)).ToString("P");
            lbl_porc_tuberia.Text = (Convert.ToDecimal(txt_tuberia.Text) / Convert.ToDecimal(lbl_met_tube.Text)).ToString("P");
            lbl_porc_incrusta.Text = (Convert.ToDecimal(txt_incrus.Text) / Convert.ToDecimal(lbl_met_incrust.Text)).ToString("P");
            lbl_porc_tapetes.Text = (Convert.ToDecimal(txt_tapetes.Text) / Convert.ToDecimal(lbl_met_tapetes.Text)).ToString("P");
            lbl_porc_gabine.Text = (Convert.ToDecimal(txt_gabine.Text) / Convert.ToDecimal(lbl_met_gabin.Text)).ToString("P");
            lbl_porc_gres.Text = (Convert.ToDecimal(txt_gres.Text) / Convert.ToDecimal(lbl_met_gres.Text)).ToString("P");
            lbl_porc_adhesivo.Text = (Convert.ToDecimal(txt_adhesivos.Text) / Convert.ToDecimal(lbl_met_adhesivos.Text)).ToString("P");
            lbl_porc_perfil.Text = (Convert.ToDecimal(txt_perfiles.Text) / Convert.ToDecimal(lbl_met_perfiles.Text)).ToString("P");
            lbl_porc_electrico.Text = (Convert.ToDecimal(txt_electrico.Text) / Convert.ToDecimal(lbl_met_electri.Text)).ToString("P");
            lbl_porc_cocinas.Text = (Convert.ToDecimal(txt_cocinas.Text) / Convert.ToDecimal(lbl_met_cocinas.Text)).ToString("P");
            lbl_porc_prodQuimi.Text = (Convert.ToDecimal(txt_prod_quim.Text) / Convert.ToDecimal(lbl_met_produQui.Text)).ToString("P");
        }


        private void txt_cemen_boqui_Leave(object sender, EventArgs e)
        {

            if (txt_cemen_boqui.Text != "")
            {
                lbl_porc_cem_boqui.Text = (Convert.ToDecimal(txt_cemen_boqui.Text) / Convert.ToDecimal(lbl_met_cemen_boq.Text)).ToString("P");   
                return;
            }
           
            lbl_porc_cem_boqui.Text = "0.00 %";


        }

        private void txt_cenefa_Leave(object sender, EventArgs e)
        {
            if (txt_cenefa.Text != "")
            {
                lbl_porc_cenefa.Text = (Convert.ToDecimal(txt_cenefa.Text) / Convert.ToDecimal(lbl_met_cenefa.Text)).ToString("P");
                return;
            }

            lbl_porc_cenefa.Text = "0.00 %";
           
        }

        private void txt_ceramica_Leave(object sender, EventArgs e)
        {
            if (txt_ceramica.Text != "")
            {
                lbl_porc_cera.Text = (Convert.ToDecimal(txt_ceramica.Text) / Convert.ToDecimal(lbl_met_cera.Text)).ToString("P");
                return;
            }

           lbl_porc_cera.Text= "0.00 %";
        }

        private void txt_griferia_Leave(object sender, EventArgs e)
        {
            if (txt_griferia.Text != "")
            {
                lbl_porc_griferia.Text = (Convert.ToDecimal(txt_griferia.Text) / Convert.ToDecimal(lbl_met_grife.Text)).ToString("P");
                return;
            }

            lbl_porc_griferia.Text = "0.00 %";
        }

        private void txt_lavaplatos_Leave(object sender, EventArgs e)
        {
            if (txt_lavaplatos.Text != "")
            {
               lbl_porc_lavapla.Text = (Convert.ToDecimal(txt_lavaplatos.Text) / Convert.ToDecimal(lbl_met_lavapla.Text)).ToString("P");
                return;
            }

            lbl_porc_lavapla.Text = "0.00 %";
        }

        private void txt_pegos_Leave(object sender, EventArgs e)
        {
            if (txt_pegos.Text != "")
            {
                lbl_porc_pego.Text = (Convert.ToDecimal(txt_pegos.Text) / Convert.ToDecimal(lbl_met_pegos.Text)).ToString("P");
                return;
            }

            lbl_porc_pego.Text = "0.00 %";
        }

        private void txt_pintura_Leave(object sender, EventArgs e)
        {
            if (txt_pintura.Text != "")
            {
                lbl_porc_pintura.Text = (Convert.ToDecimal(txt_pintura.Text) / Convert.ToDecimal(lbl_met_pint.Text)).ToString("P");
                return;
            }

            lbl_porc_pintura.Text = "0.00 %";
        }

        private void txt_porcela_Leave(object sender, EventArgs e)
        {
            if (txt_porcela.Text != "")
            {
                lbl_porc_porcela.Text = (Convert.ToDecimal(txt_porcela.Text) / Convert.ToDecimal(lbl_met_porcela.Text)).ToString("P");
                return;
            }

            lbl_porc_porcela.Text = "0.00 %";
        }

        private void txt_sanitario_Leave(object sender, EventArgs e)
        {
            if (txt_sanitario.Text != "")
            {
            lbl_porc_sanita.Text = (Convert.ToDecimal(txt_sanitario.Text) / Convert.ToDecimal(lbl_met_sanita.Text)).ToString("P");
                return;
            }

            lbl_porc_sanita.Text = "0.00 %";
        }

        private void txt_brochas_Leave(object sender, EventArgs e)
        {
            if (txt_brochas.Text != "")
            {
                lbl_porc_brochas.Text = (Convert.ToDecimal(txt_brochas.Text) / Convert.ToDecimal(lbl_met_brochass.Text)).ToString("P");
                return;
            }

            lbl_porc_brochas.Text = "0.00 %";
        }

        private void txt_lijas_Leave(object sender, EventArgs e)
        {
            if (txt_lijas.Text != "")
            {
                lbl_porc_lijas.Text= (Convert.ToDecimal(txt_lijas.Text) / Convert.ToDecimal(lbl_met_lijas.Text)).ToString("P");
                return;
            }

            lbl_porc_lijas.Text = "0.00 %";
        }

        private void txt_tuberia_Leave(object sender, EventArgs e)
        {
            if (txt_tuberia.Text != "")
            {
                lbl_porc_tuberia.Text = (Convert.ToDecimal(txt_tuberia.Text) / Convert.ToDecimal(lbl_met_tube.Text)).ToString("P");
                return;
            }

            lbl_porc_tuberia.Text = "0.00 %";
        }

        private void txt_incrus_Leave(object sender, EventArgs e)
        {
            if (txt_incrus.Text != "")
            {
               lbl_porc_incrusta.Text= (Convert.ToDecimal(txt_incrus.Text) / Convert.ToDecimal(lbl_met_incrust.Text)).ToString("P");
                return;
            }

            lbl_porc_incrusta.Text = "0.00 %";
        }

        private void txt_tapetes_Leave(object sender, EventArgs e)
        {
            if (txt_tapetes.Text != "")
            {
                lbl_porc_tapetes.Text = (Convert.ToDecimal(txt_tapetes.Text) / Convert.ToDecimal(lbl_met_tapetes.Text)).ToString("P");
                return;
            }

            lbl_porc_tapetes.Text = "0.00 %";
        }

        private void txt_gabine_Leave(object sender, EventArgs e)
        {
            if (txt_gabine.Text != "")
            {
               lbl_porc_gabine.Text = (Convert.ToDecimal(txt_gabine.Text) / Convert.ToDecimal(lbl_met_gabin.Text)).ToString("P");
                return;
            }

            lbl_porc_gabine.Text = "0.00 %";
        }

        private void txt_gres_Leave(object sender, EventArgs e)
        {
            if (txt_gres.Text != "")
            {
                lbl_porc_gres.Text = (Convert.ToDecimal(txt_gres.Text) / Convert.ToDecimal(lbl_met_gres.Text)).ToString("P");
                return;
            }

            lbl_porc_gres.Text = "0.00 %";
        }

        private void txt_adhesivos_Leave(object sender, EventArgs e)
        {
            if (txt_adhesivos.Text != "")
            {
                lbl_porc_adhesivo.Text = (Convert.ToDecimal(txt_adhesivos.Text) / Convert.ToDecimal(lbl_met_adhesivos.Text)).ToString("P");
                return;
            }

            lbl_porc_adhesivo.Text = "0.00 %";
        }

        private void txt_perfiles_Leave(object sender, EventArgs e)
        {
            if (txt_perfiles.Text != "")
            {
                lbl_porc_perfil.Text= (Convert.ToDecimal(txt_perfiles.Text) / Convert.ToDecimal(lbl_met_perfiles.Text)).ToString("P");
                return;
            }

            lbl_porc_perfil.Text = "0.00 %";
        }

        private void txt_electrico_Leave(object sender, EventArgs e)
        {
            if (txt_electrico.Text != "")
            {
              lbl_porc_electrico.Text= (Convert.ToDecimal(txt_electrico.Text) / Convert.ToDecimal(lbl_met_electri.Text)).ToString("P");
                return;
            }

            lbl_porc_electrico.Text = "0.00 %";
        }

        private void txt_cocinas_Leave(object sender, EventArgs e)
        {
            if (txt_cocinas.Text != "")
            {
                lbl_porc_cocinas.Text = (Convert.ToDecimal(txt_cocinas.Text) / Convert.ToDecimal(lbl_met_cocinas.Text)).ToString("P");
                return;
            }

            lbl_porc_cocinas.Text = "0.00 %";
        }

        private void txt_prod_quim_Leave(object sender, EventArgs e)
        {
            if (txt_prod_quim.Text != "")
            {
                lbl_porc_prodQuimi.Text= (Convert.ToDecimal(txt_prod_quim.Text) / Convert.ToDecimal(lbl_met_produQui.Text)).ToString("P");
                return;
            }

            lbl_porc_prodQuimi.Text = "0.00 %";
        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void txt_ventas_semanal_KeyPress(object sender, KeyPressEventArgs e)
        {
          
        
        
    }

        private void txt_ventas_semanal_Leave(object sender, EventArgs e)
        {
            if (txt_ventas_semanal.Text != "")
            {
                lbl_porc_cumpl_semana.Text = (Convert.ToDecimal(txt_ventas_semanal.Text) / Convert.ToDecimal(lbl_total_semana.Text)).ToString("P");
                return;
            }

            lbl_porc_cumpl_semana.Text = "0.00 %";
        }

        private void button1_Click_3(object sender, EventArgs e)
        {
            this.Dispose();
        }
     }
}
