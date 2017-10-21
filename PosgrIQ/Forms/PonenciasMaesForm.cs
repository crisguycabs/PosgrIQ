﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using GemBox.Spreadsheet;

namespace PosgrIQ
{
    public partial class PonenciasMaesForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        /// <summary>
        /// contiene los indices de las filas que coinciden con la busqueda
        /// </summary>
        public List<int> busqueda = null;

        /// <summary>
        /// indice del elemento de la busqueda que esta seleccionado
        /// </summary>
        public int iBusqueda = -1;

        #endregion

        public PonenciasMaesForm()
        {
            InitializeComponent();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarPonenciasMaesForm();
        }

        public void PonenciasMaesForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtEstudiantesMaes, dtPonenciasMaes;
                OleDbDataAdapter da;

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtEstudiantesMaes = new DataTable();
                da.Fill(dtEstudiantesMaes);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                // se pide la informacion de las ponencias de doctorado
                query = "SELECT * FROM PonenciasMaes ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtPonenciasMaes = new DataTable();
                da.Fill(dtPonenciasMaes);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                // se crea una tabla que reemplaza el codigo de colegiatura y de escuela por el nombre
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Estudiante", typeof(string));
                dt.Columns.Add("Nombre", typeof(string));
                dt.Columns.Add("Revista", typeof(string));
                dt.Columns.Add("Fecha", typeof(string));
                dt.Columns.Add("Alcance", typeof(string));

                // se llena el nuevo datatable
                for (int i = 0; i < dtPonenciasMaes.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtPonenciasMaes.Rows[i][0];

                    // nombre del estudiante
                    DataRow[] seleccionado = dtEstudiantesMaes.Select("codigo=" + dtPonenciasMaes.Rows[i][1]);
                    fila[1] = Convert.ToString(seleccionado[0][1]);

                    // titulo de la publicacion
                    fila[2] = dtPonenciasMaes.Rows[i][2];

                    // revista de la publicacion
                    fila[3] = dtPonenciasMaes.Rows[i][3];

                    // fecha de la publicacion
                    fila[4] = Convert.ToString(dtPonenciasMaes.Rows[i][4].ToString());

                    // alcance de la publicacion (Nacional o Internacional)
                    fila[5] = dtPonenciasMaes.Rows[i][5];

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridPonencias.DataSource = dt;

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridPonencias);

                dataGridPonencias.Sort(dataGridPonencias.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Ponencias de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddPonenciasMaesForm agregar = new AddPonenciasMaesForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.PonenciasMaesForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridPonencias.Rows.Count < 1) return;

            AddPonenciasMaesForm modificar = new AddPonenciasMaesForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridPonencias.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.PonenciasMaesForm_Load(sender, e);
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            // se vacia la lista de indices de la busqueda
            busqueda = new List<int>();

            if (txtSearch.Text.Length < 1) return;

            // se prepara la cadena de busqueda
            string searchValue = txtSearch.Text;

            dataGridPonencias.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                /*foreach (DataGridViewRow row in dataGridEstudiantes.Rows)
                {
                    if (row.Cells[1].Value.ToString().Contains(searchValue))
                    {
                        // row.Selected = true;
                        // dataGridEstudiantes.FirstDisplayedScrollingRowIndex = dataGridEstudiantes.SelectedRows[0].Index;
                        // break;
                        
                    }
                }*/

                for (int i = 0; i < dataGridPonencias.Rows.Count; i++)
                {
                    if (dataGridPonencias.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                    {
                        busqueda.Add(i);
                    }
                }

                // se selecciona la primera fila encontrada, si existe una
                if (busqueda.Count > 0)
                {
                    iBusqueda = 0;
                    dataGridPonencias.Rows[busqueda[0]].Selected = true;
                    dataGridPonencias.FirstDisplayedScrollingRowIndex = dataGridPonencias.SelectedRows[0].Index;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            // failsafe
            if (busqueda.Count <= 0) return;

            // se aumenta en 1 el indice de busqueda seleccionado
            iBusqueda++;
            if (iBusqueda >= busqueda.Count) iBusqueda = busqueda.Count - 1;

            dataGridPonencias.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridPonencias.FirstDisplayedScrollingRowIndex = dataGridPonencias.SelectedRows[0].Index;
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            // failsafe
            if (busqueda.Count <= 0) return;

            iBusqueda--;
            if (iBusqueda <= 0) iBusqueda = 0;

            dataGridPonencias.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridPonencias.FirstDisplayedScrollingRowIndex = dataGridPonencias.SelectedRows[0].Index;
        }

        private void dataGridEstudiantes_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
    }
}
