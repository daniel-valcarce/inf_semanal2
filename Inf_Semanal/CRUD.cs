using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace Inf_Semanal
{
   public static class CRUD
    {

        public static string string_conexion = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
         @"Data source= \\Mc01\mc01\basededatos\" +
         @"Basedatosinfsemanal.mdb";
       /* public static string string_conexion = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
         @"Data source= \\SUBGERENTE\sistema de gestion de calidad\" +
         @"Basedatosinfsemanal.mdb";*/


        public static OleDbConnection conectar() {
            OleDbConnection conexion = new OleDbConnection(string_conexion);
            return conexion;
        }


        public static System.Data.DataTable selec_meses() { 
            OleDbConnection conexion = conectar();
        OleDbCommand selec_usu = new OleDbCommand("SELECT * FROM meses", conexion);

                    conexion.Open();

                    OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();

                    DataRow row;
                    System.Data.DataTable tablameses = new System.Data.DataTable();
                    DataColumn columna;
                    columna = new DataColumn();
                    columna.ColumnName = "MES";
                    columna.DataType = System.Type.GetType("System.String");
                    tablameses.Columns.Add(columna);
                    columna = new DataColumn();
                    columna.ColumnName = "ID";
                    columna.DataType = System.Type.GetType("System.Int32");
                    tablameses.Columns.Add(columna);
                    while (lectorRegistros.Read())
                    {
                        row = tablameses.NewRow();
                        row["ID"] = lectorRegistros["id"];
                        row["MES"] = lectorRegistros["mes"].ToString();
                        tablameses.Rows.Add(row);

                    }
                    return tablameses;
        }



        public static System.Data.DataTable selec_semanas() {
            OleDbConnection conexion = conectar();
            OleDbCommand selec_usu = new OleDbCommand("SELECT * FROM semanas", conexion);

            conexion.Open();

            OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();

            DataRow row;
            System.Data.DataTable tablameses = new System.Data.DataTable();
            DataColumn columna;
            columna = new DataColumn();
            columna.ColumnName = "SEMANA";
            columna.DataType = System.Type.GetType("System.String");
            tablameses.Columns.Add(columna);
            columna = new DataColumn();
            columna.ColumnName = "ID";
            columna.DataType = System.Type.GetType("System.Int32");
            tablameses.Columns.Add(columna);
            while (lectorRegistros.Read())
            {
                row = tablameses.NewRow();
                row["ID"] = lectorRegistros["id"];
                row["SEMANA"] = lectorRegistros["semana"].ToString();
                tablameses.Rows.Add(row);

            }
            return tablameses;
        }



        public static System.Data.DataTable selec_semanas( int mes)
        {
            OleDbConnection conexion = conectar();
            OleDbCommand selec_usu = new OleDbCommand(string.Format("SELECT semanas.id, semanas.semana FROM semanas, dias_habiles2 WHERE dias_habiles2.mes={0} AND semanas.id=dias_habiles2.semana", mes), conexion);

            conexion.Open();

            OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();

            DataRow row;
            System.Data.DataTable tablameses = new System.Data.DataTable();
            DataColumn columna;
            columna = new DataColumn();
            columna.ColumnName = "SEMANA";
            columna.DataType = System.Type.GetType("System.String");
            tablameses.Columns.Add(columna);
            columna = new DataColumn();
            columna.ColumnName = "ID";
            columna.DataType = System.Type.GetType("System.Int32");
            tablameses.Columns.Add(columna);
            while (lectorRegistros.Read())
            {
                row = tablameses.NewRow();
                row["ID"] = lectorRegistros["id"];
                row["SEMANA"] = lectorRegistros["semana"].ToString();
                tablameses.Rows.Add(row);

            }
            return tablameses;
        }





       
        public static OleDbDataReader Determina_Oper(int id_user, int id_mes, int id_semana, int tipo_asesor)
        {
            OleDbConnection conexion = conectar();
            if(tipo_asesor==1){
            OleDbCommand selec_usu = new OleDbCommand(string.Format("SELECT * FROM informes_central WHERE id_user = {0} and id_mes ={1} and id_semana={2}", id_user, id_mes, id_semana), conexion);
            conexion.Open();

            OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();
            if (lectorRegistros.HasRows)
            {
                return lectorRegistros;
            }
            else
            {
                lectorRegistros = null;
                return lectorRegistros;
            }
            }
            if (tipo_asesor == 2)
            {
                OleDbCommand selec_usu = new OleDbCommand(string.Format("SELECT * FROM informes_punto where id_user = {0} and id_mes ={1} and id_semana={2}", id_user, id_mes, id_semana), conexion);
                conexion.Open();

                OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();
                if (lectorRegistros.HasRows)
                {
                    return lectorRegistros;
                }
                else
                {
                    lectorRegistros = null;
                    return lectorRegistros;
                }
            }
            else{ 
            OleDbCommand selec_usu = new OleDbCommand(string.Format("SELECT * FROM informes_externos where id_user = {0} and id_mes ={1} and id_semana={2}", id_user, id_mes, id_semana), conexion);
            conexion.Open();

            OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();
            if (lectorRegistros.HasRows)
            {
                return lectorRegistros;
            }
            else
            {
                lectorRegistros = null;
                return lectorRegistros;
            }
            }
           
        }




   

        public static int insert_infor_semanal(int id_user, int id_mes, int id_semana, int tipo_asesor, double[] valores)
        {

            OleDbConnection conexion = conectar();
            if (tipo_asesor == 1)
            {
                OleDbCommand insert_infor = new OleDbCommand(string.Format("INSERT INTO informes_central(CEMENTO_BOQUILLA,CENEFA,CERAMICA,GABINETES_ESPEJOS,GRES,GRIFERIA,LAVAPLATOS,PEGOS,PERFILES,PINTURA,PORCELANATO,PRODUCTOS_QUIMICOS,SANITARIOS,TAPETES,LIJAS,BROCHAS,ADHESIVOS,TUBERIA,ELECTRICO,COCINAS_INTREGALES,INCRUSTACIONES,VENTAS_SEMANAL,id_semana,id_user,id_mes) VALUES({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24})", valores[0], valores[1], valores[2], valores[3], valores[4], valores[5], valores[6], valores[7], valores[8], valores[9], valores[10], valores[11], valores[12], valores[13], valores[14], valores[15], valores[16], valores[17], valores[18], valores[19], valores[20], valores[21], id_semana, id_user, id_mes), conexion);
                conexion.Open();

               int lectorRegistros = insert_infor.ExecuteNonQuery();
                if (lectorRegistros>0)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            
            if (tipo_asesor == 2)
            {
                OleDbCommand selec_usu = new OleDbCommand(string.Format("INSERT INTO informes_punto(CEMENTO_BOQUILLA,CENEFA,CERAMICA,GABINETES_ESPEJOS,GRES,GRIFERIA,LAVAPLATOS,PEGOS,PERFILES,PINTURA,PORCELANATO,PRODUCTOS_QUIMICOS,SANITARIOS,TAPETES,LIJAS,BROCHAS,ADHESIVOS,TUBERIA,ELECTRICO,COCINAS_INTREGALES,INCRUSTACIONES,VENTAS_SEMANAL,id_semana,id_user,id_mes) VALUES({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24})", valores[0], valores[1], valores[2], valores[3], valores[4], valores[5], valores[6], valores[7], valores[8], valores[9], valores[10], valores[11], valores[12], valores[13], valores[14], valores[15], valores[16], valores[17], valores[18], valores[19], valores[20], valores[21], id_semana, id_user, id_mes), conexion);
                conexion.Open();

                int filas_afect = selec_usu.ExecuteNonQuery();
                if (filas_afect>0)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                OleDbCommand selec_usu = new OleDbCommand(string.Format("INSERT INTO informes_externos(CEMENTO_BOQUILLA,CENEFA,CERAMICA,GRIFERIA,LAVAPLATOS,PEGOS,PINTURA,PORCELANATO,SANITARIOS,LIJAS,BROCHAS,TUBERIA,INCRUSTACIONES,VENTAS_SEMANAL,id_semana,id_user,id_mes) VALUES({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16})", valores[0], valores[1], valores[2], valores[5], valores[6], valores[7], valores[9], valores[10], valores[12], valores[14], valores[15], valores[17], valores[20], valores[21], id_semana, id_user, id_mes), conexion);
                conexion.Open();

                int filas_afect = selec_usu.ExecuteNonQuery();
                if (filas_afect>0)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
                
            }

        }

        public static System.Data.DataTable selec_usuario() {

            OleDbConnection conexion = conectar();
            OleDbCommand selec_usu = new OleDbCommand("SELECT * FROM acceso", conexion);

            conexion.Open();

            OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();

            DataRow row;
            System.Data.DataTable tabla = new System.Data.DataTable();
            DataColumn columna;
            columna = new DataColumn();
            columna.ColumnName = "NOMBRES";
            columna.DataType = System.Type.GetType("System.String");
            tabla.Columns.Add(columna);
            columna = new DataColumn();
            columna.ColumnName = "ID";
            columna.DataType = System.Type.GetType("System.Int32");
            tabla.Columns.Add(columna);
            while (lectorRegistros.Read())
            {
                row = tabla.NewRow();
                row["ID"] = lectorRegistros["id"];
                row["NOMBRES"] = lectorRegistros["nombres"].ToString() +" "+ lectorRegistros["apellidos"].ToString();
                tabla.Rows.Add(row);

            }
            return tabla;
        
        }


        public static System.Data.DataTable tipo_asesor() {
            OleDbConnection conexion = conectar();
            OleDbCommand selec_usu = new OleDbCommand("SELECT * FROM tipo_asesor", conexion);

            conexion.Open();

            OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();

            DataRow row;
            System.Data.DataTable tabla = new System.Data.DataTable();
            DataColumn columna;
            columna = new DataColumn();
            columna.ColumnName = "TIPO";
            columna.DataType = System.Type.GetType("System.String");
            tabla.Columns.Add(columna);
            columna = new DataColumn();
            columna.ColumnName = "ID";
            columna.DataType = System.Type.GetType("System.Int32");
            tabla.Columns.Add(columna);
            while (lectorRegistros.Read())
            {
                row = tabla.NewRow();
                row["ID"] = lectorRegistros["id"];
                row["TIPO"] = lectorRegistros["tipo"].ToString();
                tabla.Rows.Add(row);

            }
            return tabla;

        }


        public static OleDbDataReader obtener_Metas(int tipo_asesor)
        {
            OleDbConnection conexion = conectar();
           
            if(tipo_asesor==1){
                OleDbCommand selec_usu = new OleDbCommand("SELECT * FROM metas_central", conexion);
                conexion.Open();
                OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();
                    return lectorRegistros;
            }
            if (tipo_asesor == 2)
            {
                OleDbCommand selec_usu = new OleDbCommand("SELECT * FROM metas_punto", conexion);
                conexion.Open();
                OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();
                return lectorRegistros;
            }
            else {
                OleDbCommand selec_usu = new OleDbCommand("SELECT * FROM metas_externo", conexion);
                conexion.Open();
                OleDbDataReader lectorRegistros = selec_usu.ExecuteReader();
                return lectorRegistros;
            }
            
         }


        public static int update_informe(int id_user, int id_mes, int id_semana, int tipo_asesor, double[] valores) {
            OleDbConnection conexion = conectar();
            if (tipo_asesor == 1)
            {
                OleDbCommand insert_infor = new OleDbCommand(string.Format("UPDATE informes_central SET CEMENTO_BOQUILLA= {0} ,CENEFA={1},CERAMICA={2},GABINETES_ESPEJOS={3},GRES={4},GRIFERIA={5},LAVAPLATOS={6},PEGOS={7},PERFILES={8},PINTURA={9},PORCELANATO={10},PRODUCTOS_QUIMICOS={11},SANITARIOS={12},TAPETES={13},LIJAS={14},BROCHAS={15},ADHESIVOS={16},TUBERIA={17},ELECTRICO={18},COCINAS_INTREGALES={19},INCRUSTACIONES={20}, VENTAS_SEMANAL={21} WHERE id_semana={22} AND id_user={23} AND id_mes={24}", valores[0], valores[1], valores[2], valores[3], valores[4], valores[5], valores[6], valores[7], valores[8], valores[9], valores[10], valores[11], valores[12], valores[13], valores[14], valores[15], valores[16], valores[17], valores[18], valores[19], valores[20], valores[21], id_semana, id_user, id_mes), conexion);
                conexion.Open();

                int lectorRegistros = insert_infor.ExecuteNonQuery();
                if (lectorRegistros > 0)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }if (tipo_asesor == 2)
            {
                OleDbCommand selec_usu = new OleDbCommand(string.Format("UPDATE informes_punto SET CEMENTO_BOQUILLA={0},CENEFA={1},CERAMICA={2},GABINETES_ESPEJOS={3},GRES={4},GRIFERIA={5},LAVAPLATOS={6},PEGOS={7},PERFILES={8},PINTURA={9},PORCELANATO={10},PRODUCTOS_QUIMICOS={11},SANITARIOS={12},TAPETES={13},LIJAS={14},BROCHAS={15},ADHESIVOS={16},TUBERIA={17},ELECTRICO={18},COCINAS_INTREGALES={19},INCRUSTACIONES={20}, VENTAS_SEMANAL={21} WHERE id_semana={22} AND id_user={23} AND id_mes={24}", valores[0], valores[1], valores[2], valores[3], valores[4], valores[5], valores[6], valores[7], valores[8], valores[9], valores[10], valores[11], valores[12], valores[13], valores[14], valores[15], valores[16], valores[17], valores[18], valores[19], valores[20], valores[21], id_semana, id_user, id_mes), conexion);
                conexion.Open();

                int lectorRegistros = selec_usu.ExecuteNonQuery();
                if (lectorRegistros>0)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                OleDbCommand selec_usu = new OleDbCommand(string.Format("UPDATE informes_externos SET CEMENTO_BOQUILLA={0},CENEFA={1},CERAMICA={2},GRIFERIA={3},LAVAPLATOS={4},PEGOS={5},PINTURA={6},PORCELANATO={7},SANITARIOS={8},LIJAS={9},BROCHAS={10},TUBERIA={11},INCRUSTACIONES={12}, VENTAS_SEMANAL={13} WHERE id_semana={14} AND id_user={15} AND id_mes={16}", valores[0], valores[1], valores[2], valores[5], valores[6], valores[7], valores[9], valores[10], valores[12], valores[14], valores[15], valores[17], valores[20], valores[21], id_semana, id_user, id_mes), conexion);
                conexion.Open();

                int lectorRegistros = selec_usu.ExecuteNonQuery();
                if (lectorRegistros>0)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }

            }


        }


        public static OleDbDataReader dias_habiles(int id_mes, int id_semana) {
            OleDbConnection conexion = conectar();

                OleDbCommand porc_cumpli = new OleDbCommand(string.Format("SELECT * FROM dias_habiles2 WHERE mes={0} and semana ={1}",id_mes, id_semana), conexion);
                conexion.Open();
                OleDbDataReader lectorRegistros = porc_cumpli.ExecuteReader();
                return lectorRegistros;
            
        }




        public static OleDbDataReader tipo_de_asesor_seleccionado(int id_user)
        {
            OleDbConnection conexion = conectar();
           
            OleDbCommand porc_cumpli = new OleDbCommand(string.Format("SELECT * FROM acceso WHERE id={0}", id_user), conexion);
            conexion.Open();
            OleDbDataReader lectorRegistros = porc_cumpli.ExecuteReader();
            return lectorRegistros;

        }



    }


}
