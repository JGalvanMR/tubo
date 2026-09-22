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
    public partial class form_tubo3 : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1, reader2;
        SqlCommand cmnd2;
        SqlCommand cmnd1;

        DataTable tarimas = new DataTable();
        DataTable tarimas2 = new DataTable();

        decimal mtot = 0;
        string mnom = "";

        string check = "";
        string tx1 = "";
        
        public form_tubo3(string Text1, string chkpla, string lin)
        {
            InitializeComponent();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            this.check = chkpla;
            this.tx1 = Text1;

            tarimas.Columns.Add("clave", typeof(string));
            tarimas.Columns.Add("producto", typeof(string));
            tarimas.Columns.Add("cajas", typeof(decimal));
            tarimas.Columns.Add("tarima", typeof(decimal));
            tarimas.Columns.Add("tomar", typeof(bool));

            tarimas2.Columns.Add("clave", typeof(string));
            tarimas2.Columns.Add("producto", typeof(string));
            tarimas2.Columns.Add("cajas", typeof(decimal));
            tarimas2.Columns.Add("tarima", typeof(decimal));
            tarimas2.Columns.Add("tomar", typeof(bool));

            thisConnection.Open();
            if (chkpla == "1")
            {
                if (Convert.ToInt32(Text1) > 120000)
                {
                    label3.Text = traerlin(lin);
                    btnTodos.Enabled = true;
                }
                else
                {
                    label3.Text = "Materia Prima";
                    btnTodos.Enabled = false;
                }

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT D.cve_prod, D.num_cajas, D.tarima, P.prod_nombre FROM tb_det_eti_final D, tb_cat_producto P WHERE D.cve_prod = P.prod_clave AND folio = '" + Text1 + "'";
                reader1 = cmnd1.ExecuteReader();
                DataRow dr;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        dr = tarimas.NewRow();
                        dr["clave"] = reader1.GetValue(0).ToString().Trim();
                        dr["producto"] = reader1.GetValue(3).ToString().Trim();
                        dr["cajas"] = reader1.GetDecimal(1);
                        dr["tarima"] = reader1.GetDecimal(2).ToString("0");
                        dr["tomar"] = false;
                        tarimas.Rows.Add(dr);
                        mtot = mtot + reader1.GetDecimal(1);
                    }
                }

                label4.Text = mtot.ToString("###,###,##0.00");

                DataView dw = tarimas.DefaultView;
                dw.Sort = "clave, tarima";
                tarimas = dw.ToTable();

                foreach (DataRow r in tarimas.Rows)
                {
                    dtgProductos.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), Convert.ToBoolean(r[4].ToString()));
                }
            }
            else
            {
                string a = Text1;
                cmnd1 = thisConnection.CreateCommand();
                if (Convert.ToInt32(Text1) > 120000)
                {
                    label3.Text = traerlin(lin);
                    btnTodos.Enabled = true;
                }
                else
                {
                    label3.Text = "Materia Prima";
                    btnTodos.Enabled = false;
                }

                if (Convert.ToInt32(Text1) < 150000 && Convert.ToInt32(Text1) > 120000)
                {
                    cmnd1.CommandText = "SELECT H.hrp_recibo, T.prod_clave, T.prod_nombre, T.etiqueta, T.tarima" +
                        " FROM tb_hist_recepcion H, tb_det_trazabilidad T" +
                        " WHERE H.hrp_recibo = T.recibo AND H.prod_clave = T.prod_clave" +
                        " AND H.hrp_tipo_recepcion = 'MP' AND H.hrp_recibo = '" + Text1 + "'" +
                        " AND NOT H.hrp_recibo IN" +
                        " (SELECT folio FROM tb_det_tubo" +
                        " WHERE tipo = 'MP' AND folio = '" + Text1 + "' AND cve_prod = H.prod_clave AND tarima = T.tarima)";

                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            DataRow r = tarimas.NewRow();
                            r["clave"] = reader1.GetValue(1).ToString().Trim();
                            r["producto"] = reader1.GetValue(2).ToString().Trim();
                            r["cajas"] = reader1.GetDecimal(3);
                            r["tarima"] = reader1.GetDecimal(4).ToString("0");
                            r["tomar"] = false;
                            tarimas.Rows.Add(r);

                            mtot = mtot + reader1.GetDecimal(3);
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();

                    DataView dw = tarimas.DefaultView;
                    dw.Sort = "clave, tarima";
                    tarimas = dw.ToTable();

                    foreach (DataRow r in tarimas.Rows)
                    {
                        dtgProductos.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), Convert.ToBoolean(r[4].ToString()));
                    }

                    label4.Text = mtot.ToString("###,###,##0.00");
                        
                }
                if (Convert.ToInt32(Text1) > 150000 && Convert.ToInt32(Text1) > 120000)
                {
                    cmnd1.CommandText = "SELECT H.hrp_recibo, T.prod_clave, T.prod_nombre, T.etiqueta, T.tarima" +
                        " FROM tb_hist_recepcion H, tb_det_trazabilidad T" +
                        " WHERE H.hrp_recibo = T.recibo AND H.prod_clave = T.prod_clave" +
                        " AND H.hrp_tipo_recepcion = 'PTC' AND H.hrp_recibo = '" + Text1 + "'" +
                        " AND NOT H.hrp_recibo IN" +
                        " (SELECT folio FROM tb_det_tubo" +
                        " WHERE tipo = 'PTC' AND folio = '" + Text1 + "' AND cve_prod = H.prod_clave AND tarima = T.tarima)";

                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            DataRow r = tarimas.NewRow();
                            r["clave"] = reader1.GetValue(1).ToString().Trim();
                            r["producto"] = reader1.GetValue(2).ToString().Trim();
                            r["cajas"] = reader1.GetDecimal(3);
                            r["tarima"] = reader1.GetDecimal(4).ToString("0");
                            r["tomar"] = false;
                            tarimas.Rows.Add(r);

                            mtot = mtot + reader1.GetDecimal(3);
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();

                    cmnd1.CommandText = "SELECT H.hrp_recibo, T.prod_clave, T.prod_nombre, T.etiqueta, T.tarima" +
                        " FROM tb_hist_recepcion H, tb_det_trazabilidad T" +
                        " WHERE H.hrp_recibo = T.recibo AND H.prod_clave = T.prod_clave" +
                        " AND H.hrp_tipo_recepcion = 'PTC' AND H.hrp_recibo = '" + Text1 + "'" +
                        " AND H.hrp_recibo IN" +
                        " (SELECT folio FROM tb_det_tubo" +
                        " WHERE tipo = 'PTC' AND folio = '" + Text1 + "' AND cve_prod = H.prod_clave AND tarima = T.tarima)";

                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            mtot = mtot + reader1.GetDecimal(3);
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();

                    DataView dw = tarimas.DefaultView;
                    dw.Sort = "clave, tarima";
                    tarimas = dw.ToTable();

                    foreach (DataRow r in tarimas.Rows)
                    {
                        dtgProductos.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), Convert.ToBoolean(r[4].ToString()));
                    }

                    label4.Text = mtot.ToString("###,###,##0.00");
                }
                if (Convert.ToInt32(Text1) < 150000 && Convert.ToInt32(Text1) < 120000)
                {

                    //cmnd1.CommandText = "SELECT H.hrp_recibo, H.prod_clave, P.prod_nombre, H.hrp_num_unidades, T.cantidad, T.tipo FROM tb_hist_recepcion H, tb_cat_producto P, tb_det_tubo T" +
                    //    " WHERE H.hrp_recibo = T.folio AND H.prod_clave = P.prod_clave AND H.hrp_tipo_recepcion = T.tipo AND T.tipo = 'MP'" +
                    //    " AND H.hrp_recibo = '" + Text1 + "'";
                    cmnd1.CommandText = "SELECT H.hrp_recibo, H.prod_clave, P.prod_nombre, H.hrp_num_unidades, T.cantidad, T.tipo FROM tb_hist_recepcion H, tb_cat_producto P, tb_det_tubo T" +
                        " WHERE H.hrp_recibo = T.folio AND H.prod_clave = P.prod_clave AND H.hrp_tipo_recepcion = T.tipo AND T.tipo = 'MP'" +
                        " AND H.hrp_recibo = '" + Text1 + "'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            if (reader1.GetDecimal(3) < reader1.GetDecimal(4))
                            {
                                MessageBox.Show("Del recibo ya se enfriaron todas las cajas", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                DataRow r = tarimas.NewRow();
                                r["clave"] = reader1.GetValue(1).ToString().Trim();
                                r["producto"] = reader1.GetValue(2).ToString().Trim();
                                r["cajas"] = reader1.GetDecimal(3) - reader1.GetDecimal(4);
                                tarimas.Rows.Add(r);

                                mtot = mtot + reader1.GetDecimal(3);
                            }
                        }

                    }
                    else
                    {
                        reader1.Close();
                        reader1.Dispose();

                        cmnd1.CommandText = "SELECT H.hrp_recibo, H.prod_clave, P.prod_nombre, H.hrp_num_unidades FROM tb_hist_recepcion H, tb_cat_producto P" +
                        " WHERE H.prod_clave = P.prod_clave AND H.hrp_recibo = '" + Text1 + "'" +
                        " AND H.hrp_tipo_recepcion = 'MP' AND NOT H.hrp_recibo IN (SELECT folio FROM tb_det_tubo WHERE tipo = 'MP')";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            while (reader1.Read())
                            {
                                DataRow r = tarimas.NewRow();
                                r["clave"] = reader1.GetValue(1).ToString().Trim();
                                r["producto"] = reader1.GetValue(2).ToString().Trim();
                                r["cajas"] = reader1.GetDecimal(3);
                                tarimas.Rows.Add(r);

                                mtot = mtot + reader1.GetDecimal(3);
                            }
                        }
                        reader1.Close();
                        reader1.Dispose();
                    }

                    DataView dw = tarimas.DefaultView;
                    dw.Sort = "clave";
                    tarimas = dw.ToTable();

                    foreach (DataRow r in tarimas.Rows)
                    {
                        dtgProductos.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), "0", false);
                    }

                    label4.Text = mtot.ToString("###,###,##0.00");
                }
                
            }
                
            thisConnection.Close();

        }

        public string traerlin(string cve)
        {
            string cad = "";
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "SELECT lin_nombre FROM tb_cat_linea WHERE lin_clave = '" + cve + "' ORDER BY lin_clave";
            reader2 = cmnd2.ExecuteReader();
            if (reader2.HasRows)
            {
                reader2.Read();
                cad = reader2.GetValue(0).ToString().Trim();
            }
            reader2.Close();
            reader2.Dispose();

            return cad;
        }

        public class datosval
        {
            private DataTable _elementos;
            public DataTable elementos
            {
                get { return _elementos; }
                set { _elementos = value; }
            }

            private decimal _tot;
            public decimal tot
            {
                get { return _tot; }
                set { _tot = value; }
            }

            private string _hora;
            public string hora
            {
                get { return _hora; }
                set { _hora = value; }
            }
        }
        public class SharedData
        {
            public static datosval Polino;
        }

        private void dtgProductos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                if ((bool)dtgProductos.Rows[dtgProductos.CurrentRow.Index].Cells[4].EditedFormattedValue == true)
                {
                    tarimas.Rows[dtgProductos.CurrentRow.Index][e.ColumnIndex] = true;
                }
                else
                {
                    tarimas.Rows[dtgProductos.CurrentRow.Index][e.ColumnIndex] = false;
                }
                //MessageBox.Show(dtgProductos.Rows[dtgProductos.CurrentRow.Index].Cells[4].EditedFormattedValue.ToString());
                
            }
        }

        private void btnAplicar_Click(object sender, EventArgs e)
        {
            if (tarimas.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos en el listado", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int counter = 0;
            foreach (DataRow rw in tarimas.Select("tomar = 'true'"))
            {
                counter++;
            }
            if (counter == 0)
            {
                MessageBox.Show("No hay datos seleccionados para aplicar", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            mtot = 0;
            int entra = 0;
            string hr = "";
            foreach (DataRow r in tarimas.Select("tomar = 'true'"))
            {

                if (this.check == "1")
                {
                    tarimas2.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), ((r[3].ToString() != "") ? Convert.ToDecimal(r[3].ToString()) : 0), Convert.ToBoolean(r[4].ToString()));
                    //if (entra == 0)
                    //{
                        thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT hora FROM tb_det_eti_final WHERE folio = '" + this.tx1.PadLeft(6, '0') + "' AND cve_prod = '" + r[0].ToString() + "' ORDER BY cve_prod, tarima";
                        reader1 = cmnd1.ExecuteReader();
                        while (reader1.Read())
                        {
                            hr = reader1.GetValue(0).ToString();
                        }
                        thisConnection.Close();        
                    //}
                    
                }
                else
                {
                    tarimas2.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), ((r[3].ToString() != "") ? Convert.ToDecimal(r[3].ToString()) : 0), Convert.ToBoolean(r[4].ToString()));
                }
                mtot = mtot + Convert.ToDecimal(r[2].ToString());

            }

            if (this.check == "0")
            {
                if (Convert.ToInt32(tx1) > 120000)
                {
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT rpt_hora FROM tb_mstr_recepcion_pt WHERE rpt_recibo = '" + this.tx1 + "' ORDER BY rpt_recibo";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        hr = reader1.GetValue(0).ToString();
                    }
                    thisConnection.Close();
                }
                else
                {
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT rmp_hora FROM tb_mstr_recepcion_mp WHERE rmp_recibo = '" + this.tx1 + "' ORDER BY rmp_recibo";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        hr = reader1.GetValue(0).ToString();
                    }
                    thisConnection.Close();
                }
            }

            datosval passdata = new datosval();
            passdata.elementos = tarimas2;
            passdata.hora = hr;
            passdata.tot = mtot;
            SharedData.Polino = passdata;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnTodos_Click(object sender, EventArgs e)
        {
            if (tarimas.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos en el listado", "TUBO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            foreach(DataRow r in tarimas.Rows)
            {
                r[4] = true;
            }

            mtot = 0;
            int entra = 0;
            string hr = "";
            foreach (DataRow r in tarimas.Select("tomar = 'true'"))
            {

                if (this.check == "1")
                {
                    tarimas2.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), ((r[3].ToString() != "") ? Convert.ToDecimal(r[3].ToString()) : 0), Convert.ToBoolean(r[4].ToString()));
                    //if (entra == 0)
                    //{
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT hora FROM tb_det_eti_final WHERE folio = '" + this.tx1.PadLeft(6, '0') + "' AND cve_prod = '" + r[0].ToString() + "' ORDER BY cve_prod, tarima";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        hr = reader1.GetValue(0).ToString();
                    }
                    thisConnection.Close();
                    //}

                }
                else
                {
                    tarimas2.Rows.Add(r[0].ToString(), r[1].ToString(), r[2].ToString(), ((r[3].ToString() != "") ? Convert.ToDecimal(r[3].ToString()) : 0), Convert.ToBoolean(r[4].ToString()));
                }
                mtot = mtot + Convert.ToDecimal(r[2].ToString());

            }

            if (this.check == "0")
            {
                if (Convert.ToInt32(tx1) > 120000)
                {
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT rpt_hora FROM tb_mstr_recepcion_pt WHERE rpt_recibo = '" + this.tx1 + "' ORDER BY rpt_recibo";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        hr = reader1.GetValue(0).ToString();
                    }
                    thisConnection.Close();
                }
            }

            datosval passdata = new datosval();
            passdata.elementos = tarimas2;
            passdata.hora = hr;
            passdata.tot = mtot;
            SharedData.Polino = passdata;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
