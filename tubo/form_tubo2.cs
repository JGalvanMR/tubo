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
using System.Text.RegularExpressions;

namespace tubo
{
    public partial class form_tubo2 : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1, reader2, reader3;
        SqlCommand cmnd2, cmnd3;
        SqlCommand cmnd1;

        DataTable detalle = new DataTable();
        DataTable tarimas = new DataTable();

        public string opcion = "";

        public string lin_cve = "";

        List<string> dts = new List<string>();

        public form_tubo2(string opc, List<string> dat)
        {
            InitializeComponent();

            textBox1.Focus();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            this.opcion = opc;
            this.dts = dat;

            detalle.Columns.Add("folio", typeof(string));
            detalle.Columns.Add("prod", typeof(string));
            detalle.Columns.Add("nombre", typeof(string));
            detalle.Columns.Add("cant", typeof(decimal));
            detalle.Columns.Add("tarima", typeof(decimal));

            tarimas.Columns.Add("clave", typeof(string));
            tarimas.Columns.Add("producto", typeof(string));
            tarimas.Columns.Add("cajas", typeof(decimal));
            tarimas.Columns.Add("tariam", typeof(decimal));
            tarimas.Columns.Add("tomar", typeof(bool));

            opciones(this.opcion);
        }

        public void opciones(string opci)
        {
            switch (opci)
            { 
                case "A":
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox13.Enabled = false;
                    textBox14.Enabled = false;
                    break;
                case "C":
                    pbxNuevo.Visible = false;
                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    textBox1.Enabled = false;
                    pictureBox2.Visible = false;
                    chkplanta.Enabled = false;
                    dtgDatos.ReadOnly = true;
                    textBox3.Enabled = false;
                    textBox4.Enabled = false;
                    textBox5.Enabled = false;
                    textBox6.Enabled = false;
                    textBox7.Enabled = false;
                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox13.Enabled = false;
                    textBox14.Enabled = false;

                    comboBox1.Text = this.dts[4].ToString();
                    comboBox2.Text = this.dts[1].ToString();
                    textBox1.Text = this.dts[2].ToString();
                    lblNombreProd.Text = this.dts[3].ToString();
                    textBox3.Text = this.dts[6].ToString();
                    textBox4.Text = this.dts[5].ToString();
                    textBox5.Text = this.dts[7].ToString();
                    textBox6.Text = this.dts[8].ToString();
                    textBox7.Text = this.dts[9].ToString();
                    textBox8.Text = this.dts[10].ToString();
                    textBox9.Text = this.dts[11].ToString();
                    textBox10.Text = this.dts[12].ToString();
                    textBox11.Text = this.dts[13].ToString();

                    string t = "";
                    string he = "";
                    string pr = "";
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT phagua, cloro, tipo, hora_ent FROM tb_mstr_tubo WHERE folio = '" + this.dts[2].ToString() + "'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        textBox13.Text = reader1.GetDecimal(0).ToString("000.00");
                        textBox14.Text = reader1.GetDecimal(1).ToString("000.00");
                        t = reader1.GetValue(2).ToString().Trim();
                        he = reader1.GetValue(3).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();

                    cmnd1.CommandText = "SELECT T.folio, T.cve_prod, P.prod_nombre, T.cantidad, T.tarima FROM tb_det_tubo T, tb_cat_producto P" +
                        " WHERE T.cve_prod = P .prod_clave AND T.folio = '" + this.dts[2].ToString() + "' AND T.tipo = '" + t + "' AND T.hora_ent = '" + he + "'";
                    reader1 = cmnd1.ExecuteReader();
                    DataRow r;
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            r = detalle.NewRow();
                            r["folio"] = reader1.GetValue(0).ToString().Trim();
                            r["prod"] = reader1.GetValue(1).ToString().Trim();
                            r["nombre"] = reader1.GetValue(2).ToString().Trim();
                            r["cant"] = reader1.GetValue(3).ToString().Trim();
                            r["tarima"] = reader1.GetValue(4).ToString().Trim();
                            pr = reader1.GetValue(1).ToString().Trim();
                            detalle.Rows.Add(r);
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();

                    if (t == "MP")
                    {
                        cmnd1.CommandText = "SELECT M.prov_clave, M.rch_clave, M.tbl_clave, P.prov_nombre, R.rch_nombre, T.tbl_nombre" +
                            " FROM tb_mstr_tubo TU, tb_mstr_recepcion_mp M, tb_cat_proveedor P, tb_cat_ranchos R, tb_cat_tablas T" +
                            " WHERE TU.folio = M.rmp_recibo" +
                            " AND M.prov_clave = P.prov_clave" +
                            " AND M.prov_clave = R.prov_clave" +
                            " AND M.rch_clave = R.rch_clave" +
                            " AND M.prov_clave = T.prov_clave" +
                            " AND M.rch_clave = T.rch_clave" +
                            " AND M.tbl_clave = T.tbl_clave" +
                            " AND M.rmp_recibo = '" + this.dts[2].ToString() + "'";
                        reader1 = cmnd1.ExecuteReader();
                        while (reader1.Read())
                        {
                            label18.Text = reader1.GetValue(3).ToString().Trim();
                            label19.Text = reader1.GetValue(4).ToString().Trim();
                            label21.Text = reader1.GetValue(5).ToString().Trim();
                        }
                        reader1.Close();
                        reader1.Dispose();
                    }
                    else
                    {
                        cmnd1.CommandText = "SELECT TR.prov_nombre, TR.rch_nombre, TR.tbl_nombre FROM tb_det_trazabilidad TR, tb_mstr_tubo T" +
                            " WHERE T.folio = TR.recibo AND TR.recibo = '" + this.dts[2].ToString() + "' AND prod_clave = '" + pr + "'";
                        reader1 = cmnd1.ExecuteReader();
                        while (reader1.Read())
                        {
                            label18.Text = reader1.GetValue(0).ToString().Trim();
                            label19.Text = reader1.GetValue(1).ToString().Trim();
                            label21.Text = reader1.GetValue(2).ToString().Trim();
                        }
                        reader1.Close();
                        reader1.Dispose();
                    }
                    thisConnection.Close();

                    foreach (DataRow rw in detalle.Rows)
                    {
                        dtgDatos.Rows.Add(rw[0].ToString(), rw[1].ToString(), rw[2].ToString(), rw[3].ToString(), rw[4].ToString());
                    }

                    if (Convert.ToInt32(this.dts[11].ToString()) > 0 && Convert.ToInt32(this.dts[12].ToString()) > 0 && Convert.ToInt32(this.dts[13].ToString()) > 0)
                    {
                        MessageBox.Show("Ya se capturaró el recibo completo!!  Solo esta disponible para consulta", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        textBox9.Enabled = true;
                        textBox10.Enabled = true;
                        textBox11.Enabled = true;
                        textBox13.Enabled = true;
                        textBox14.Enabled = true;
                        pbxNuevo.Visible = true;
                    }
                    break;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            string A = "";
            if (e.KeyChar == (char)Keys.Enter || e.KeyChar == (char)Keys.Tab)
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                A = textBox1.Text.PadLeft(6, '0');

                if (chkplanta.Checked == true)
                {
                    
                    cmnd1.CommandText = "SELECT ordp_folio FROM tb_mstr_ordenes_prod WHERE ordp_folio = '" + A + "' ORDER BY ordp_folio";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows == false)
                    {
                        MessageBox.Show("No existe el recibo de producción", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        thisConnection.Close();
                        pictureBox2.Enabled = false;
                        pictureBox2.Visible = false;
                        return;
                    }
                    reader1.Close();
                    reader1.Dispose();
                    thisConnection.Close();
                    pictureBox2.Enabled = true;
                    pictureBox2.Visible = true;
                    return;
                }

                if(Convert.ToInt32(A) < 150000)
                    cmnd1.CommandText = "SELECT hrp_recibo, lin_clave FROM tb_hist_recepcion WHERE hrp_tipo_recepcion = 'MP' AND hrp_recibo = '" + A + "' ORDER BY hrp_tipo_recepcion, hrp_recibo, lin_clave, prod_clave";
                else
                    cmnd1.CommandText = "SELECT hrp_recibo, lin_clave FROM tb_hist_recepcion WHERE hrp_tipo_recepcion = 'PTC' AND hrp_recibo = '" + A + "' ORDER BY hrp_tipo_recepcion, hrp_recibo, lin_clave, prod_clave";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows == false)
                {
                    MessageBox.Show("No existe el recibo", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    thisConnection.Close();
                    pictureBox2.Enabled = true;
                    pictureBox2.Visible = true;
                    return;
                }
                else
                {
                    reader1.Read();
                    lin_cve = reader1.GetValue(1).ToString().Trim();
                }
                reader1.Close();
                reader1.Dispose();
                thisConnection.Close();

                pictureBox2.Enabled = true;
                pictureBox2.Visible = true;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            string chk = (chkplanta.Checked == true) ? "1" : "0";
            form_tubo3 dlg = new form_tubo3(textBox1.Text, chk, lin_cve);
            dlg.ShowDialog();

            if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                tarimas.Clear();
                tarimas = form_tubo3.SharedData.Polino.elementos;
                DataView dw = tarimas.DefaultView;
                dw.Sort = "producto";
                tarimas = dw.ToTable();

                foreach (DataRow r in tarimas.Rows)
                {
                    dtgDatos.Rows.Add(textBox1.Text, r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString());
                }

                lblNombreProd.Text = tarimas.Rows[0][1].ToString();
                if (form_tubo3.SharedData.Polino.hora != "")
                {
                    textBox3.Text = form_tubo3.SharedData.Polino.hora;
                    textBox3.Enabled = false;
                }
                    


                textBox4.Text = form_tubo3.SharedData.Polino.tot.ToString("###,###,##0.00");
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pbxNuevo_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Se va a grabar la información esta seguro??", "TUBO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                thisConnection.Open();
                if (opcion == "A")
                {
                    if (textBox1.Text == "")
                    {
                        MessageBox.Show("No se ha capturado el folio", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (textBox5.Text == "")
                    {
                        MessageBox.Show("No se ha capturado la hora de entrada al tubo", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (comboBox1.SelectedIndex == -1)
                    {
                        MessageBox.Show("No se ha capturado el operador", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (comboBox2.SelectedIndex == -1)
                    {
                        MessageBox.Show("No se ha capturado el lugar de refrigeración", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (textBox3.Text == "")
                    {
                        MessageBox.Show("No se ha capturado la hora de llegada", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (textBox4.Text == "")
                    {
                        MessageBox.Show("No se ha capturado la cantidad de cajas", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (textBox7.Text == "")
                    {
                        MessageBox.Show("No se ha capturado la temperatura de entrada al tubo", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (textBox8.Text == "" && comboBox2.SelectedItem.ToString() != "FRESCO")
                    {
                        MessageBox.Show("No se ha capturado la temperatura inicial del agua", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (DataRow r in tarimas.Rows)
                    {
                        detalle.Rows.Add(textBox1.Text, r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString());
                    }

                    //VALIDACION DE HORA LLEGO Y HORA ENTRO
                    if (validahora(textBox3.Text) == false)
                    {
                        MessageBox.Show("La hora de llegada no es válida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (validahora(textBox5.Text) == false)
                    {
                        MessageBox.Show("La hora de entrada no es válida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    string hr_ent = "";
                    hr_ent = textBox5.Text;
                    string mfol = "";
                    string mti = "";
                    string mtar = "";

                    foreach (DataRow r in detalle.Rows)
                    {
                        mfol = r[0].ToString();
                        mti = (Convert.ToInt32(mfol) > 170000) ? "PTC" : "MP";
                        mtar = r[4].ToString();

                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "SELECT folio FROM tb_det_tubo WHERE folio = '" + mfol + "' AND tipo = '" + mti + "'" +
                        //    " AND hora_ent = '" + hr_ent + "' AND tarima = '" + mtar + "'";
                        //reader1 = cmnd1.ExecuteReader();
                        //if (reader1.HasRows)
                        //{
                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = "INSERT INTO" +
                                " tb_det_tubo(folio, tipo, cve_prod, cantidad, hora_ent, tarima)" +
                                " VALUES('" + mfol + "', '" + mti + "', '" + r[1].ToString() + "', '" + r[3].ToString() + "', '" + hr_ent + "', '" + mtar + "')";
                            cmnd3.ExecuteNonQuery();
                            cmnd3.Dispose();
                        //}
                        //reader1.Close();
                        //reader1.Dispose();

                        cmnd1.CommandText = "SELECT folio FROM tb_mstr_tubo WHERE folio = '" + mfol + "' AND hora_ent = '" + hr_ent + "'";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows == false)
                        {
                            string mhr = "";
                            cmnd2 = thisConnection.CreateCommand();
                            if (Convert.ToInt32(mfol) > 170000)
                            {
                                cmnd2.CommandText = "SELECT rpt_hora FROM tb_mstr_recepcion_pt WHERE rpt_recibo = '" + mfol + "' ORDER BY rpt_recibo DESC";
                                reader2 = cmnd2.ExecuteReader();
                                if (reader2.HasRows)
                                {
                                    reader2.Read();
                                    mhr = (reader2.GetValue(0).ToString().Trim().Length > 5) ? reader2.GetValue(0).ToString().Trim().Substring(0,5) : "00:00";
                                }
                                reader2.Close();
                                reader2.Dispose();
                                cmnd2.Dispose();
                            }
                            else
                            {
                                mhr = textBox3.Text;
                            }
                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = "INSERT INTO" +
                                " tb_mstr_tubo(fecha, lugar, folio, tipo, cve_prod, nom_operador, hora_llego, cantidad, hora_ent, temp_ent, temp_iniagua," +
                                " temp_sal, temp_finagua, duracion, phagua, cloro)" +
                                " VALUES('" + dtpFecha.Text + "', '" + traelugar(comboBox2.SelectedIndex) + "', '" + mfol + "', '" + mti + "'," +
                                " '" + r[1].ToString() + "', '" + comboBox1.SelectedItem.ToString() + "', '" + mhr + "', '" + r[3].ToString() + "'," +
                                " '" + hr_ent + "', '" + textBox7.Text + "', '" + textBox8.Text + "', '0', '0', '0', '0', '0')";
                            cmnd3.ExecuteNonQuery();
                            cmnd3.Dispose();
                        }
                        else
                        {
                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = "UPDATE tb_mstr_tubo SET cantidad = cantidad + '" + Convert.ToDecimal(r[3].ToString()) + "'" +
                                " WHERE folio = '" + mfol + "' AND hora_ent = '" + hr_ent + "'";
                            cmnd3.ExecuteNonQuery();
                            cmnd3.Dispose();
                        }
                        reader1.Close();
                        reader1.Dispose();
                    }

                    MessageBox.Show("Datos guardados con éxito", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    textBox1.Text = "";
                    lblNombreProd.Text = "-";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    tarimas.Clear();
                    detalle.Clear();
                }
                else
                {
                    if (textBox9.Text == "")
                    {
                        MessageBox.Show("No se ha capturado la temperatura de salida al tubo", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (textBox10.Text == "")
                    {
                        MessageBox.Show("No se ha capturado la temperatura final del agua", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (textBox11.Text == "")
                    {
                        MessageBox.Show("No se ha capturado el tiempo de duración en el tubo", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    //Validacion sólo números
                    if (validavalor(textBox9.Text) == false)
                    {
                        MessageBox.Show("El valor de la temperatura de salida no es numérico", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (validavalor(textBox10.Text) == false)
                    {
                        MessageBox.Show("El valor de la temperatura de final del agua no es numérico", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (validavalor(textBox11.Text) == false)
                    {
                        MessageBox.Show("El valor del tiempo de duración no es numérico", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }


                    //validacion mayor a cero
                    if (Convert.ToDecimal(textBox9.Text) <= 0)
                    {
                        MessageBox.Show("El valor de la temperatura de salida debe ser mayor a cero", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (Convert.ToDecimal(textBox10.Text) <= 0)
                    {
                        MessageBox.Show("El valor de la temperatura de final del agua debe ser mayor a cero", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (Convert.ToDecimal(textBox11.Text) <= 0)
                    {
                        MessageBox.Show("El valor del tiempo de duración debe ser mayor a cero", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "UPDATE tb_mstr_tubo SET temp_sal = '" + textBox9.Text + "', temp_finagua = '" + textBox10.Text + "'," +
                        " duracion = '" + textBox11.Text + "', phagua = '" + textBox13.Text + "', cloro = '" + textBox14.Text + "' WHERE folio = '" + textBox1.Text + "' AND hora_ent = '" + textBox5.Text + "'";
                    cmnd1.ExecuteNonQuery();

                }
                thisConnection.Close();
                DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }

        public string traelugar(int indx)
        {
            string cad = "";
            switch (indx)
            { 
                case 0:
                    cad = "TUBO1";
                    break;
                case 1:
                    cad = "TUBO2";
                    break;
                case 2:
                    cad = "ENFRIA";
                    break;
                case 3:
                    cad = "HIDROC";
                    break;
                case 4:
                    cad = "HIDROG";
                    break;
                case 5:
                    cad = "ENHIE";
                    break;
            }
            return cad;
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (textBox5.Text == "")
                {
                    MessageBox.Show("No ha ingresado una hora de entrada", "SIPGAB", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Regex checktime = new Regex(@"^(?:0?[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$");
                bool valido = false;

                if (checktime.IsMatch(textBox5.Text) == true)
                {
                    valido = true;
                }
                else
                {
                    valido = false;
                }

                if (valido == false)
                {
                    MessageBox.Show("La hora introducida no es válida", "SIPGAB", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                //DateTime hora1 = DateTime.Parse(textBox3.Text);
                //DateTime hora2 = DateTime.Parse(textBox5.Text);
                //DateTime hora3 = DateTime.Parse("23:59");

                //TimeSpan Span = hora2.Subtract(hora1);
                //TimeSpan Span2;
                //if (Span.Hours < 0)
                //{
                //    Span2 = hora3.Subtract(hora1);
                //    DateTime Total = DateTime.Parse(Span2.Hours + ":" + Span2.Minutes);
                //    Total.AddHours(hora2.Hour);
                //    Total.AddMinutes(hora2.Minute);
                //    textBox6.Text = Total.ToString("HH:mm");
                //}
                //else
                //{
                //    DateTime Total = DateTime.Parse(Span.Hours + ":" + Span.Minutes);
                //    textBox6.Text = Total.ToString("HH:mm");
                //}
                //MessageBox.Show(calculatiempo(textBox3.Text, textBox5.Text));
                
                textBox6.Text = calculatiempo(textBox3.Text, textBox5.Text);
                
            }
        }

        public string calculatiempo(string TI, string TF)
        {
            TimeSpan tmp;
            string t = "";
            int h;
            int m;
            if (Convert.ToDateTime(TF).Hour == 0)
            {
                h = 23 - Convert.ToDateTime(TI).Hour;
                m = (60 - Convert.ToDateTime(TI).Minute) + Convert.ToDateTime(TF).Minute;
                t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
            }
            else
            {
                tmp = Convert.ToDateTime(TF) - Convert.ToDateTime(TI);
                if (tmp.Hours < 0)
                {
                    h = 23 - Convert.ToDateTime(TI).Hour;
                    m = (60 - Convert.ToDateTime(TI).Minute);
                    //h = h + Convert.ToDateTime(TF).Hour;
                    t = Convert.ToString(h).PadLeft(2, '0') + ":" + m.ToString().PadLeft(2, '0') + ":00";
                    t = Convert.ToDateTime(t).AddHours(Convert.ToDateTime(TF).Hour).ToString();
                    t = Convert.ToDateTime(t).AddMinutes(Convert.ToDateTime(TF).Minute).ToString("HH:mm");
                }
                else
                {
                    //h = tmp.Minutes / 3600;
                    //m = (h - Convert.ToInt32(h)) * 60;
                    t = tmp.ToString();
                }
            }
            return t;

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

        public bool validavalor(string val)
        {
            bool valido = false;
            decimal valor = 0;
            try
            {
                valor = Convert.ToDecimal(val);
                return true;
            }
            catch (Exception ex)
            {
                return valido;
            }
        }
    }
}
