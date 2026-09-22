using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using SISEMP;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Collections;

namespace tubo
{
    public partial class Form1 : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1, reader2;
        SqlCommand cmnd2;
        SqlCommand cmnd1;

        DataTable temperaturas = new DataTable();
        DataTable temperaturasexcel = new DataTable();
        DataTable temperaturasexceldet = new DataTable();
        DataTable detallesexcel = new DataTable();

        public List<string> datos = new List<string>();

        public string opcion = "";
        int contador = 0;

        public Form1()
        {
            InitializeComponent();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            label5.Text = DateTime.Now.ToString();

            temperaturas.Columns.Add("fecha", typeof(string));
            temperaturas.Columns.Add("lugar", typeof(string));
            temperaturas.Columns.Add("folio", typeof(string));
            temperaturas.Columns.Add("producto", typeof(string));
            temperaturas.Columns.Add("operador", typeof(string));
            temperaturas.Columns.Add("cantidad", typeof(string));
            temperaturas.Columns.Add("hrllego", typeof(string));
            temperaturas.Columns.Add("hrentro", typeof(string));
            temperaturas.Columns.Add("tiempo", typeof(string));
            temperaturas.Columns.Add("tempentro", typeof(string));
            temperaturas.Columns.Add("tempiniagua", typeof(string));
            temperaturas.Columns.Add("tempsalida", typeof(string));
            temperaturas.Columns.Add("tmpfinagua", typeof(string));
            temperaturas.Columns.Add("duracion", typeof(string));
            temperaturas.Columns.Add("lugar2", typeof(string));

            temperaturasexcel.Columns.Add("fecha", typeof(string));
            temperaturasexcel.Columns.Add("folio", typeof(string));
            temperaturasexcel.Columns.Add("lugar", typeof(string));
            temperaturasexcel.Columns.Add("cveprod", typeof(string));
            temperaturasexcel.Columns.Add("producto", typeof(string));
            temperaturasexcel.Columns.Add("operador", typeof(string));
            temperaturasexcel.Columns.Add("hrllego", typeof(string));
            temperaturasexcel.Columns.Add("cantidad", typeof(string));
            temperaturasexcel.Columns.Add("hrentro", typeof(string));
            temperaturasexcel.Columns.Add("tiempo", typeof(string));
            temperaturasexcel.Columns.Add("duracion", typeof(string));
            temperaturasexcel.Columns.Add("hrsalio", typeof(string));
            temperaturasexcel.Columns.Add("tempentro", typeof(string));
            temperaturasexcel.Columns.Add("tempiniagua", typeof(string));
            temperaturasexcel.Columns.Add("tempsalida", typeof(string));
            temperaturasexcel.Columns.Add("tempfinagua", typeof(string));
            temperaturasexcel.Columns.Add("phagua", typeof(string));
            temperaturasexcel.Columns.Add("cloro", typeof(string));

            temperaturasexceldet.Columns.Add("fecha", typeof(string));
            temperaturasexceldet.Columns.Add("folio", typeof(string));
            temperaturasexceldet.Columns.Add("lugar", typeof(string));
            temperaturasexceldet.Columns.Add("cveprod", typeof(string));
            temperaturasexceldet.Columns.Add("producto", typeof(string));
            temperaturasexceldet.Columns.Add("operador", typeof(string));
            temperaturasexceldet.Columns.Add("hrllego", typeof(string));
            temperaturasexceldet.Columns.Add("cantidad", typeof(string));
            temperaturasexceldet.Columns.Add("hrentro", typeof(string));
            temperaturasexceldet.Columns.Add("tiempo", typeof(string));
            temperaturasexceldet.Columns.Add("duracion", typeof(string));
            temperaturasexceldet.Columns.Add("hrsalio", typeof(string));
            temperaturasexceldet.Columns.Add("tempentro", typeof(string));
            temperaturasexceldet.Columns.Add("tempiniagua", typeof(string));
            temperaturasexceldet.Columns.Add("tempsalida", typeof(string));
            temperaturasexceldet.Columns.Add("tempfinagua", typeof(string));
            temperaturasexceldet.Columns.Add("phagua", typeof(string));
            temperaturasexceldet.Columns.Add("cloro", typeof(string));

            detallesexcel.Columns.Add("folio", typeof(string));
            detallesexcel.Columns.Add("cve_prod", typeof(string));
            detallesexcel.Columns.Add("producto", typeof(string));
            detallesexcel.Columns.Add("tarima", typeof(string));
            detallesexcel.Columns.Add("cantidad", typeof(string));
            detallesexcel.Columns.Add("tipo", typeof(string));
            detallesexcel.Columns.Add("hora_ent", typeof(string));

            cargargrid();
            
        }

        private void pbxFiltro_Click(object sender, EventArgs e)
        {
            try
            {
                temperaturas.Clear();
                dtgTiempos.Rows.Clear();
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT fecha, lugar, folio, tipo, cve_prod, nom_operador, hora_llego, cantidad, hora_ent, temp_ent, temp_iniagua, temp_sal, temp_finagua, duracion, phagua, cloro " +
                    "FROM tb_mstr_tubo WHERE fecha = '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' ORDER BY folio, hora_llego";
                reader1 = cmnd1.ExecuteReader();
                DataRow dr;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        if (reader1.GetValue(2).ToString().Trim() == "224624")
                        { 
                        }
                        dr = temperaturas.NewRow();
                        dr["fecha"] = reader1.GetValue(0).ToString().Trim();
                        dr["lugar"] = reader1.GetValue(1).ToString().Trim();
                        dr["folio"] = reader1.GetValue(2).ToString().Trim();
                        dr["producto"] = reader1.GetValue(4).ToString().Trim();
                        dr["operador"] = reader1.GetValue(5).ToString().Trim();
                        dr["cantidad"] = reader1.GetValue(7).ToString().Trim();
                        dr["hrllego"] = reader1.GetValue(6).ToString().Trim();
                        dr["hrentro"] = reader1.GetValue(8).ToString().Trim();
                        dr["tiempo"] = calculatiempo(reader1.GetValue(6).ToString().Trim(), reader1.GetValue(8).ToString().Trim());
                        dr["tempentro"] = reader1.GetValue(9).ToString().Trim();
                        dr["tempiniagua"] = reader1.GetValue(10).ToString().Trim();
                        dr["tempsalida"] = reader1.GetValue(11).ToString().Trim();
                        dr["tmpfinagua"] = reader1.GetValue(12).ToString().Trim();
                        dr["duracion"] = reader1.GetValue(13).ToString().Trim();
                        dr["lugar"] = traelugar2(reader1.GetValue(1).ToString().Trim());
                        temperaturas.Rows.Add(dr);
                    }
                }
                thisConnection.Close();

                foreach (DataRow rw in temperaturas.Rows)
                {
                    dtgTiempos.Rows.Add(Convert.ToDateTime(rw[0].ToString()).ToString("yyyy-MM-dd"), rw[1].ToString(), rw[2].ToString(), rw[3].ToString(), rw[4].ToString(), rw[5].ToString(), rw[6].ToString(), rw[7].ToString(), rw[8].ToString(), rw[9].ToString(), rw[10].ToString(), rw[11].ToString(), rw[12].ToString(), rw[13].ToString());
                }
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }      
        }

        public string calculatiempo(string TI, string TF)
        {
            TimeSpan tmp;
            string t = "";
            int h = 0;
            int m = 0;
            if (Convert.ToDateTime(TF).Hour == 0)
            {
                if (Convert.ToDateTime(TI).Hour != 0)
                {
                    h = 24 - Convert.ToDateTime(TI).Hour;
                    m = (60 - Convert.ToDateTime(TI).Minute) + Convert.ToDateTime(TF).Minute;
                    t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                }
                else
                {
                    m = (Convert.ToDateTime(TF).Minute) - Convert.ToDateTime(TI).Minute;
                    t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                }
            }
            else
            {
                if (Convert.ToDateTime(TI) < Convert.ToDateTime(TF))
                {
                    tmp = Convert.ToDateTime(TF) - Convert.ToDateTime(TI);
                    t = tmp.ToString();
                }
                else if (Convert.ToDateTime(TI) == Convert.ToDateTime(TF))
                {
                    t = "00:00:00";
                }
                else if (Convert.ToDateTime(TF) < Convert.ToDateTime(TI))
                {
                    int t_i = Convert.ToDateTime(TI).Hour;
                    int t_f = Convert.ToDateTime(TF).Hour;

                    int t_i_m = Convert.ToDateTime(TI).Minute;
                    int t_f_m = Convert.ToDateTime(TF).Minute;

                    h = 24 - t_i + t_f;

                    if (t_f_m < t_i_m)
                    {
                        h--;
                        m = 60 - t_i_m + t_f_m;
                        t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                    }
                    else
                    {
                        //m = (Convert.ToDateTime(TF).Minute) + (60 - Convert.ToDateTime(TI).Minute);
                        //m = t_f_m + (60 - t_i_m);
                        //t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                        m = t_f_m - t_i_m;
                        t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                    }

                    
                }
                else
                {
                    //tmp = (Convert.ToDateTime(TF).AddHours(-24)) + Convert.ToDateTime(TI);

                    int t_i = Convert.ToDateTime(TI).Hour;
                    int t_f = Convert.ToDateTime(TF).Hour;

                    int t_i_m = Convert.ToDateTime(TI).Minute;
                    int t_f_m = Convert.ToDateTime(TF).Minute;

                    h = 24 - t_i + t_f;
                    //m = (Convert.ToDateTime(TF).Minute) + (60 - Convert.ToDateTime(TI).Minute);
                    m = t_f_m + (60 - t_i_m);
                    //t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                    t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                }
                //h = tmp.Minutes / 3600;
                //m = (h - Convert.ToInt32(h)) * 60;
                
            }
            return t;

        }

        public string traelugar2(string lg)
        {
            string cad = "";
            switch (lg)
            { 
                case "TUBO1":
                    cad = "TUBO1";
                    break;
                case "TUBO2":
                    cad = "TUBO2";
                    break;
                case "TUBO3":
                    cad = "TUBO3";
                    break;
                case "ENFRIA":
                    cad = "ENFRIADOR AIRE FORZADO (COLIFLOR)";
                    break;
                case "HIDROC":
                    cad = "HIDROCOOLER";
                    break;
                case "HIDROG":
                    cad = "HIDROCOOLER GRANDE";
                    break;
                case "ENHIE":
                    cad = "ENHIELADORA";
                    break;
            }
            return cad;
        }

        private void pbxNuevo_Click(object sender, EventArgs e)
        {
            opcion = "A";

            form_tubo2 dlg = new form_tubo2(opcion, datos);
            dlg.ShowDialog();
            if (dlg.DialogResult == DialogResult.OK)
            {
                pbxFiltro_Click(null, null);
            }
        }

        private void dtgTiempos_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                datos.Clear();
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][0].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][1].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][2].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][3].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][4].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][5].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][6].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][7].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][8].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][9].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][10].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][11].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][12].ToString());
                datos.Add(temperaturas.Rows[dtgTiempos.CurrentRow.Index][13].ToString());

                opcion = "C";
                form_tubo2 dlg = new form_tubo2(opcion, datos);
                dlg.ShowDialog();

                if (dlg.DialogResult == DialogResult.OK)
                {
                    pbxFiltro_Click(null, null);
                }
            }
        }

        private void pbxSalir_Click(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "SELECT TOP 1 inicio_sesion, usu_login FROM tb_cat_historial_dia where nombre_maquina = '" + Environment.MachineName + "' and sistema = 'SIPGAB' ORDER BY inicio_sesion desc";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    Utilerias.Class1.Inicio_sesion = reader2.GetSqlDateTime(0).Value;
                    Utilerias.Class1.Usu_login = reader2.GetSqlString(1).ToString();
                    Utilerias.Class1.Nombre_equipo = Environment.MachineName;
                }
                reader2.Close();

                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "update tb_cat_historial_dia set formulario = ' ' where nombre_maquina ='" + Utilerias.Class1.Nombre_equipo + "' and usu_login = '" + Utilerias.Class1.Usu_login + "' and inicio_sesion = '" + Utilerias.Class1.Inicio_sesion.ToString("s") + "' and sistema = 'SIPGAB'";
                reader2 = cmnd2.ExecuteReader();
                reader2.Close();
                thisConnection.Close();
                Application.Exit();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }      
        }

        private void pbxExcel_Click(object sender, EventArgs e)
        {
            if (Convert.ToDateTime(dtpFechaInicio.Text) > Convert.ToDateTime(dtpFechaFin.Text))
            {
                MessageBox.Show("Rango de fechas inválido", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                temperaturasexcel.Clear();
                temperaturasexceldet.Clear();
                detallesexcel.Clear();
                dtgExcel.Rows.Clear();

                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT fecha, folio, lugar, a.cve_prod, nom_operador, hora_llego, cantidad, hora_ent, duracion, temp_ent, temp_iniagua, temp_sal, temp_finagua, phagua, cloro, b.prod_nombre " +
                    "FROM tb_mstr_tubo a, tb_cat_producto b WHERE  fecha BETWEEN '" + Convert.ToDateTime(dtpFechaInicio.Text).ToShortDateString() + "' AND '" + Convert.ToDateTime(dtpFechaFin.Text).ToShortDateString() +
                    "' and a.cve_prod = b.prod_clave ORDER BY hora_llego";
                reader1 = cmnd1.ExecuteReader();
                DataRow dr;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        if (reader1.GetValue(3).ToString().Trim() == "09025")
                        { 
                        }
                        dr = temperaturasexcel.NewRow();
                        dr["fecha"] = Convert.ToDateTime(reader1.GetValue(0).ToString().Trim()).ToString("yyyy-MM-dd");
                        dr["folio"] = reader1.GetValue(1).ToString().Trim();
                        dr["lugar"] = reader1.GetValue(2).ToString().Trim();
                        dr["cveprod"] = reader1.GetValue(3).ToString().Trim();
                        dr["producto"] = reader1.GetValue(15).ToString().Trim();
                        dr["operador"] = reader1.GetValue(4).ToString().Trim();
                        dr["hrllego"] = reader1.GetValue(5).ToString().Trim();
                        dr["cantidad"] = reader1.GetValue(6).ToString().Trim();
                        dr["hrentro"] = reader1.GetValue(7).ToString().Trim();
                        dr["tiempo"] = calculatiempo(reader1.GetValue(5).ToString().Trim(), reader1.GetValue(7).ToString().Trim());
                        dr["duracion"] = reader1.GetValue(8).ToString().Trim();
                        dr["hrsalio"] = calculosalida(reader1.GetValue(7).ToString().Trim(), reader1.GetValue(8).ToString().Trim());
                        dr["tempentro"] = reader1.GetValue(9).ToString().Trim();
                        dr["tempiniagua"] = reader1.GetValue(10).ToString().Trim();
                        dr["tempsalida"] = reader1.GetValue(11).ToString().Trim();
                        dr["tempfinagua"] = reader1.GetValue(12).ToString().Trim();
                        dr["phagua"] = reader1.GetValue(13).ToString().Trim();
                        dr["cloro"] = reader1.GetValue(14).ToString().Trim();
                        temperaturasexcel.Rows.Add(dr);
                    }
                }
                reader1.Close();
                reader1.Dispose();

                DataSet set = new DataSet();
                SqlDataAdapter adapter1;

                if (chkDetalle.Checked == true)
                {
                    adapter1 = new SqlDataAdapter("SELECT RTRIM(D.folio) AS folio, RTRIM(D.tarima) AS tarima, RTRIM(D.cantidad) AS cantidad, RTRIM(D.cve_prod) AS cve_prod," +
                        " RTRIM(P.prod_nombre) AS prod_nombre, RTRIM(D.tipo) AS tipo, RTRIM(D.hora_ent) AS hora_ent FROM tb_mstr_tubo T, tb_det_tubo D, tb_cat_producto P" +
                        " WHERE T.folio = D.folio AND D.cve_prod = P.prod_clave AND T.hora_ent = D.hora_ent" +
                        " AND T.fecha BETWEEN '" + Convert.ToDateTime(dtpFechaInicio.Text).ToShortDateString() + "' AND '" + Convert.ToDateTime(dtpFechaFin.Text).ToShortDateString() + "' ORDER BY T.hora_ent, D.cve_prod, D.tarima", thisConnection);
                    adapter1.Fill(set, "Detalle");
                    detallesexcel = set.Tables["Detalle"];
                    //cmnd1.CommandText = 
                    //DataRow dr2;
                    //if (reader1.HasRows)
                    //{
                    //    while (reader1.Read())
                    //    {
                    //        dr2 = detallesexcel.NewRow();
                    //        dr2["folio"] = reader1.GetValue(0).ToString().Trim();
                    //        dr2["cve_prod"] = reader1.GetValue(3).ToString().Trim();
                    //        dr2["producto"] = reader1.GetValue(4).ToString().Trim();
                    //        dr2["tarima"] = reader1.GetValue(1).ToString().Trim();
                    //        dr2["cantidad"] = reader1.GetValue(2).ToString().Trim();
                    //        dr2["tipo"] = reader1.GetValue(5).ToString().Trim();
                    //        dr2["hora_ent"] = reader1.GetValue(6).ToString().Trim();
                    //        detallesexcel.Rows.Add(dr2);
                    //    }
                    //}
                    //reader1.Close();
                    //reader1.Dispose();
                }
                

                thisConnection.Close();


                foreach (DataRow r_1 in temperaturasexcel.Rows)
                {
                    dtgExcel.Rows.Add(r_1["fecha"].ToString(), r_1["folio"].ToString(),
                        r_1["lugar"].ToString(), r_1["cveprod"].ToString(),
                        r_1["producto"].ToString(), r_1["operador"].ToString(),
                        r_1["hrllego"].ToString(), r_1["cantidad"].ToString(),
                        r_1["hrentro"].ToString(), r_1["tiempo"].ToString(),
                        r_1["duracion"].ToString(), r_1["hrsalio"].ToString(),
                        r_1["tempentro"].ToString(), r_1["tempiniagua"].ToString(),
                        r_1["tempsalida"].ToString(), r_1["tempfinagua"].ToString(),
                        r_1["phagua"].ToString(), r_1["cloro"].ToString());
                    if (chkDetalle.Checked == true)
                    {
                        foreach (DataRow rr in detallesexcel.Select("folio = '" + r_1["folio"].ToString() + "' AND hora_ent = '" + r_1["hrentro"].ToString() + "'"))
                        {
                            dtgExcel.Rows.Add(
                            "",
                            "",
                            "",
                            rr["cve_prod"].ToString(),
                            rr["prod_nombre"].ToString(),
                            "Tar: " + rr["tarima"].ToString(),
                            "",
                            rr["cantidad"].ToString(),
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "");
                        }
                    }
                }


                //foreach (DataRow rw in temperaturas.Rows)
                //{
                //    dtgTiempos.Rows.Add(Convert.ToDateTime(rw[0].ToString()).ToString("yyyy-MM-dd"), rw[1].ToString(), rw[2].ToString(), rw[3].ToString(), rw[4].ToString(), rw[5].ToString(), rw[6].ToString(), rw[7].ToString(), rw[8].ToString(), rw[9].ToString(), rw[10].ToString(), rw[11].ToString(), rw[12].ToString(), rw[13].ToString());
                //}

                label8.Visible = true;
                label8.Update();
                progressBar1.Visible = true;
                
                //SaveFileDialog fichero = new SaveFileDialog();
                //fichero.Filter = "Excel 97 (*.xls)|*.xls";
                //if (fichero.ShowDialog() == DialogResult.OK)
                //{
                Microsoft.Office.Interop.Excel.Application aplicacion;
                Microsoft.Office.Interop.Excel.Workbook libro;
                Microsoft.Office.Interop.Excel.Worksheet hoja;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libro = aplicacion.Workbooks.Add();
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

                Microsoft.Office.Interop.Excel.Range r;
                hoja.Cells[1, 1] = "Comercializador GAB, S.A. de C.V.";
                hoja.Range[hoja.Cells[1, 1], hoja.Cells[1, 18]].Merge();
                hoja.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                r = hoja.get_Range("A1", "R1");
                r.Font.Bold = true;
                r.Font.Size = 16;

                hoja.Cells[3, 1] = "VERIFICACION DE PRODUCTO DE ENFRIAMIENTO";
                hoja.Range[hoja.Cells[3, 1], hoja.Cells[3, 18]].Merge();
                hoja.Cells[3, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                r = hoja.get_Range("A3", "R3");
                r.Font.Bold = true;
                r.Font.Size = 12;

                hoja.Cells[5, 4] = "Del " + dtpFechaInicio.Text + " al " + dtpFechaFin.Text;
                r = hoja.Range[hoja.Cells[5, 4], hoja.Cells[5, 4]];
                r.Font.Bold = true;

                hoja.Cells[7, 1] = "Fecha";
                hoja.Cells[7, 2] = "Folio";
                hoja.Cells[7, 3] = "Lugar";
                hoja.Cells[7, 4] = "Clave Prod";
                hoja.Cells[7, 5] = "Producto";
                hoja.Cells[7, 6] = "Operador";
                hoja.Cells[7, 7] = "Hr. Llegada";
                hoja.Cells[7, 8] = "Cantidad";
                hoja.Cells[7, 9] = "Hr. Entrada";
                hoja.Cells[7, 10] = "Tiempo Espera";
                hoja.Cells[7, 11] = "Duración";
                hoja.Cells[7, 12] = "Hr. Salida";
                hoja.Cells[7, 13] = "Temp. Entrada";
                hoja.Cells[7, 14] = "Temp. Inicial Agua";
                hoja.Cells[7, 15] = "Temp. Salida";
                hoja.Cells[7, 16] = "Temp. Final Agua";
                hoja.Cells[7, 17] = "Ph Agua 7.5 Max";
                hoja.Cells[7, 18] = "Cloro Libre 2 ppm Min";
                r = hoja.Range[hoja.Cells[7, 1], hoja.Cells[7, 18]];
                r.Font.Bold = true;

                Cursor.Current = Cursors.WaitCursor;
                int ult = 9;

                dtgExcel.SelectAll();
                dtgExcel.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
                Clipboard.SetDataObject(dtgExcel.GetClipboardContent().GetText());
                dtgExcel.ClearSelection();

                hoja.Cells[8, 1].Select();
                hoja.Paste();
                    
                //if (chkDetalle.Checked == true)
                //{
                //    DataRow row;
                //    for (int i = 0; i < temperaturasexcel.Rows.Count; i++)
                //    {
                //        row = temperaturasexceldet.NewRow();
                //        row["fecha"] = temperaturasexcel.Rows[i][0].ToString();
                //        row["folio"] = temperaturasexcel.Rows[i][1].ToString();
                //        row["lugar"] = temperaturasexcel.Rows[i][2].ToString();
                //        row["cveprod"] = "";
                //        row["producto"] = temperaturasexcel.Rows[i][4].ToString();
                //        row["operador"] = temperaturasexcel.Rows[i][5].ToString();
                //        row["hrllego"] = temperaturasexcel.Rows[i][6].ToString();
                //        row["cantidad"] = temperaturasexcel.Rows[i][7].ToString();
                //        row["hrentro"] = temperaturasexcel.Rows[i][8].ToString();
                //        row["tiempo"] = temperaturasexcel.Rows[i][9].ToString();
                //        row["duracion"] = temperaturasexcel.Rows[i][10].ToString();
                //        row["hrsalio"] = temperaturasexcel.Rows[i][11].ToString();
                //        row["tempentro"] = temperaturasexcel.Rows[i][12].ToString();
                //        row["tempiniagua"] = temperaturasexcel.Rows[i][13].ToString();
                //        row["tempsalida"] = temperaturasexcel.Rows[i][14].ToString();
                //        row["tempfinagua"] = temperaturasexcel.Rows[i][15].ToString();
                //        row["phagua"] = temperaturasexcel.Rows[i][16].ToString();
                //        row["cloro"] = temperaturasexcel.Rows[i][17].ToString();
                //        temperaturasexceldet.Rows.Add(row);
                            
                            
                //        foreach(DataRow rr in detallesexcel.Select("folio = '" + temperaturasexcel.Rows[i][1].ToString() + "' AND hora_ent = '" + temperaturasexcel.Rows[i][8].ToString() + "'"))
                //        {
                //            row = temperaturasexceldet.NewRow();
                //            row["fecha"] = "";
                //            row["folio"] = "";
                //            row["lugar"] = "";
                //            row["cveprod"] = rr["cve_prod"].ToString();
                //            row["producto"] = rr["producto"].ToString();
                //            row["operador"] = "Tar: " + rr["tarima"].ToString();
                //            row["hrllego"] = "";
                //            row["cantidad"] = rr["cantidad"].ToString();
                //            row["hrentro"] = "";
                //            row["tiempo"] = "";
                //            row["duracion"] = "";
                //            row["hrsalio"] = "";
                //            row["tempentro"] = "";
                //            row["tempiniagua"] = "";
                //            row["tempsalida"] = "";
                //            row["tempfinagua"] = "";
                //            row["phagua"] = "";
                //            row["cloro"] = "";
                //            temperaturasexceldet.Rows.Add(row);
                //        }
                //    }

                //    progressBar1.Maximum = temperaturasexceldet.Rows.Count;

                //    for (int i = 0; i < temperaturasexceldet.Rows.Count; i++)
                //    {
                //        for (int j = 0; j < temperaturasexceldet.Columns.Count; j++)
                //        {
                //            hoja.Cells[i + 9, j + 1] = temperaturasexceldet.Rows[i][j].ToString();
                //        }
                //        progressBar1.PerformStep();
                //        ult++;
                //    }

                //}
                //else
                //{
                //    progressBar1.Maximum = temperaturasexcel.Rows.Count;

                //    for (int i = 0; i < temperaturasexcel.Rows.Count; i++)
                //    {
                //        for (int j = 0; j < temperaturasexcel.Columns.Count; j++)
                //        {
                //            hoja.Cells[i + 9, j + 1] = temperaturasexcel.Rows[i][j].ToString();
                //        }
                //        progressBar1.PerformStep();
                //        ult++;
                            
                //    }

                        
                //}

                ult = ult + dtgExcel.Rows.Count;
                ult = ult + 1;
                hoja.Cells[ult, 1] = "F-100-20";
                ult = ult + 1;
                hoja.Cells[ult, 1] = "Rev.: 00";

                //hoja.Columns.AutoFit();
                //libro.SaveAs(fichero.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
                //libro.Close();
                //aplicacion.Quit();
                aplicacion.Columns.AutoFit();
                aplicacion.Rows.AutoFit();
                aplicacion.Visible = true;

                progressBar1.Minimum = 0;
                progressBar1.Value = 0;
                progressBar1.Visible = false;
                label8.Visible = false;

                //}

                Cursor.Current = Cursors.Default;
                //System.Diagnostics.Process.Start(@"" + fichero.FileName);
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                MessageBox.Show(ex.ToString());
            }
        }

        public string calculosalida(string he, string d)
        {
            string ho = "";
            ho = Convert.ToDateTime(he).AddMinutes(Convert.ToInt32(d)).ToString("HH:mm");
            return ho;
        }

        public void cargargrid()
        {
            try
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT fecha, lugar, folio, tipo, cve_prod, nom_operador, hora_llego, cantidad, hora_ent, temp_ent, temp_iniagua, temp_sal, temp_finagua, duracion, phagua, cloro " +
                    "FROM tb_mstr_tubo WHERE fecha = '" + DateTime.Now.ToShortDateString() + "' ORDER BY folio, hora_llego";
                reader1 = cmnd1.ExecuteReader();
                DataRow dr;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        dr = temperaturas.NewRow();
                        dr["fecha"] = reader1.GetValue(0).ToString().Trim();
                        dr["lugar"] = reader1.GetValue(1).ToString().Trim();
                        dr["folio"] = reader1.GetValue(2).ToString().Trim();
                        dr["producto"] = reader1.GetValue(4).ToString().Trim();
                        dr["operador"] = reader1.GetValue(5).ToString().Trim();
                        dr["cantidad"] = reader1.GetValue(7).ToString().Trim();
                        dr["hrllego"] = reader1.GetValue(6).ToString().Trim();
                        dr["hrentro"] = reader1.GetValue(8).ToString().Trim();
                        dr["tiempo"] = calculatiempo(reader1.GetValue(6).ToString().Trim(), reader1.GetValue(8).ToString().Trim());
                        dr["tempentro"] = reader1.GetValue(9).ToString().Trim();
                        dr["tempiniagua"] = reader1.GetValue(10).ToString().Trim();
                        dr["tempsalida"] = reader1.GetValue(11).ToString().Trim();
                        dr["tmpfinagua"] = reader1.GetValue(12).ToString().Trim();
                        dr["duracion"] = reader1.GetValue(13).ToString().Trim();
                        dr["lugar"] = traelugar2(reader1.GetValue(1).ToString().Trim());
                        temperaturas.Rows.Add(dr);
                    }
                }
                thisConnection.Close();

                foreach (DataRow rw in temperaturas.Rows)
                {
                    dtgTiempos.Rows.Add(Convert.ToDateTime(rw[0].ToString()).ToString("yyyy-MM-dd"), rw[1].ToString(), rw[2].ToString(), rw[3].ToString(), rw[4].ToString(), rw[5].ToString(), rw[6].ToString(), rw[7].ToString(), rw[8].ToString(), rw[9].ToString(), rw[10].ToString(), rw[11].ToString(), rw[12].ToString(), rw[13].ToString());
                }
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }      
        }

        private void btnminimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (contador == 300)
            {
                
                pbxFiltro_Click(null, null);
                contador = 0;
                lblTiempo.Text = "0";
            }
            lblTiempo.Text = contador++.ToString();
        }

        public bool validahora(string dato)
        {
            try
            {
                string hora = "";
                hora = Convert.ToDateTime(dato).ToString();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //pbxFiltro_Click(null, null);
        }

        //public string buscanombre(string clave)
        //{
 
        //}
    }
}
