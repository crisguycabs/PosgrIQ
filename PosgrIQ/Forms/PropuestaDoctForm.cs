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
using GemBox.Spreadsheet;

namespace PosgrIQ
{
    public partial class PropuestaDoctForm : Form
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

        public PropuestaDoctForm()
        {
            InitializeComponent();
        }

        public void PropuestaDoctForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtPropuestasDoct, dtProfesores, dtConceptos, dtEstudiantesDoct;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM PropuestaDoct ORDER BY codigo ASC";
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                conection.Open();
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtPropuestasDoct = new DataTable();
                da.Fill(dtPropuestasDoct);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                conection.Open();
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                // se pide la informacion de los profesores
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                conection.Open();
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                // se pide la informacion de los conceptos
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                conection.Open();
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtConceptos = new DataTable();
                da.Fill(dtConceptos);

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
                dt.Columns.Add("Codigo", typeof(int));                              // 0
                dt.Columns.Add("Estudiante", typeof(string));                       // 1
                dt.Columns.Add("Titulo", typeof(string));                           // 2
                dt.Columns.Add("Fecha Entrega", typeof(DateTime));                  // 3
                dt.Columns.Add("Calificador 1", typeof(string));                    // 4
                dt.Columns.Add("Calificador 2", typeof(string));                    // 5
                dt.Columns.Add("Calificador 3", typeof(string));                    // 6
                dt.Columns.Add("Calificador 4", typeof(string));                    // 7
                dt.Columns.Add("Concepto 1 Calificador 1", typeof(string));         // 8
                dt.Columns.Add("Concepto 1 Calificador 2", typeof(string));         // 9
                dt.Columns.Add("Concepto 1 Calificador 3", typeof(string));         // 10
                dt.Columns.Add("Concepto 1 Calificador 4", typeof(string));         // 11
                dt.Columns.Add("Fecha Entrega Correcciones", typeof(DateTime));     // 12
                dt.Columns.Add("Concepto 2 Calificador 1", typeof(string));         // 13
                dt.Columns.Add("Concepto 2 Calificador 2", typeof(string));         // 14
                dt.Columns.Add("Concepto 2 Calificador 3", typeof(string));         // 15
                dt.Columns.Add("Concepto 2 Calificador 4", typeof(string));         // 16
                dt.Columns.Add("Fecha Sustentacion", typeof(DateTime));             // 17
                dt.Columns.Add("Concepto Final", typeof(string));                   // 18

                string[] fechaArray = null;


                // se llena el nuevo datatable
                for (int i = 0; i < dtPropuestasDoct.Rows.Count; i++)
                {

                    try
                    {
                        DataRow fila = dt.NewRow();

                        // codigo del estudiante
                        fila[0] = dtPropuestasDoct.Rows[i][0];

                        // nombre del estudiante
                        DataRow[] seleccionado = dtEstudiantesDoct.Select("codigo=" + dtPropuestasDoct.Rows[i][1]);
                        fila[1] = Convert.ToString(seleccionado[0][1]);

                        // titulo de la propuesta
                        fila[2] = dtPropuestasDoct.Rows[i][2];

                        // fecha de entrega
                        fechaArray = dtPropuestasDoct.Rows[i][8].ToString().Split('/');
                        if (fechaArray.Length > 1) fila[3] = new DateTime(Convert.ToInt16(fechaArray[2]), Convert.ToInt16(fechaArray[1]), Convert.ToInt16(fechaArray[0]));
                        else fila[3] = new DateTime();
                        // fila[3] = Convert.ToString(dtPropuestasDoct.Rows[i][8].ToString());

                        // calificador 1
                        fila[4] = dtProfesores.Select("codigo=" + dtPropuestasDoct.Rows[i][4].ToString())[0][1];

                        // calificador 2
                        fila[5] = dtProfesores.Select("codigo=" + dtPropuestasDoct.Rows[i][5].ToString())[0][1];

                        // calificador 3
                        fila[6] = dtProfesores.Select("codigo=" + dtPropuestasDoct.Rows[i][6].ToString())[0][1];

                        // calificador 4, opcional
                        if (dtPropuestasDoct.Rows[i][7].ToString() != "")
                        {
                            // hay un calificador 4
                            fila[7] = dtProfesores.Select("codigo=" + dtPropuestasDoct.Rows[i][7].ToString())[0][1];
                        }
                        else
                        {
                            fila[7] = "";
                        }

                        // concepto 1 calificador 1
                        fila[8] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][9].ToString())[0][1];

                        // concepto 1 calificador 2
                        fila[9] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][10].ToString())[0][1];

                        // concepto 1 calificador 3
                        fila[10] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][11].ToString())[0][1];

                        // concepto 1 calificador 4
                        fila[11] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][12].ToString())[0][1];

                        // fecha de entrega de correcciones
                        fechaArray = dtPropuestasDoct.Rows[i][17].ToString().Split('/');
                        if (fechaArray.Length > 1) fila[12] = new DateTime(Convert.ToInt16(fechaArray[2]), Convert.ToInt16(fechaArray[1]), Convert.ToInt16(fechaArray[0]));
                        else fila[12] = new DateTime();
                        // fila[12] = dtPropuestasDoct.Rows[i][17];

                        // concepto 2 calificador 1
                        fila[13] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][18].ToString())[0][1];

                        // concepto 2 calificador 2
                        fila[14] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][19].ToString())[0][1];

                        // concepto 2 calificador 3
                        fila[15] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][20].ToString())[0][1];

                        // concepto 2 calificador 4
                        fila[16] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][21].ToString())[0][1];

                        // fecha sustentacion
                        fechaArray = dtPropuestasDoct.Rows[i][26].ToString().Split('/');
                        if (fechaArray.Length > 1) fila[17] = new DateTime(Convert.ToInt16(fechaArray[2]), Convert.ToInt16(fechaArray[1]), Convert.ToInt16(fechaArray[0]));
                        else fila[17] = new DateTime();
                        // fila[17] = dtPropuestasDoct.Rows[i][26];

                        // concepto sustentacion
                        fila[18] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][27].ToString())[0][1];

                        dt.Rows.Add(fila);
                    }
                    catch
                    {
                        int a = 1;
                        a++;
                    }
                }



                // se enlaza el datatable con el datagrid
                dataGridPropuestas.DataSource = dt;

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridPropuestas);

                dataGridPropuestas.Sort(dataGridPropuestas.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarPropuestasDoctForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddPropuestaDoctForm agregar = new AddPropuestaDoctForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.PropuestaDoctForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridPropuestas.Rows.Count < 1) return;

            AddPropuestaDoctForm modificar = new AddPropuestaDoctForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridPropuestas.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.PropuestaDoctForm_Load(sender, e);
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            int codigo = Convert.ToInt32(dataGridPropuestas.SelectedRows[0].Cells[0].Value);
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dtPropuestaDoct;
                OleDbDataAdapter da;

                // se pide la informacion de la propuesta de doct
                query = "SELECT * FROM PropuestaDoct WHERE codigo=" + codigo.ToString() + " ORDER BY codigo ASC";
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                conection.Open();
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtPropuestaDoct = new DataTable();
                da.Fill(dtPropuestaDoct);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                try
                {
                    switch (cmbVer.SelectedIndex)
                    {
                        case 0: // propuesta
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][3]));
                            break;
                        case 1: // concepto 1 calificador 1
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][13]));
                            break;
                        case 2: // concepto 1 calificador 2
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][14]));
                            break;
                        case 3: // concepto 1 calificador 3
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][15]));
                            break;
                        case 4: // concepto 1 calificador 4
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][16]));
                            break;
                        case 5: // concepto 2 calificador 1
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][22]));
                            break;
                        case 6: // concepto 2 calificador 2
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][23]));
                            break;
                        case 7: // concepto 2 calificador 3
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][24]));
                            break;
                        case 8: // concepto 2 calificador 4
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][25]));
                            break;
                        case 9: // sustentacion
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtPropuestaDoct.Rows[0][28]));
                            break;
                    }
                }
                catch
                {
                    MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            // se vacia la lista de indices de la busqueda
            busqueda = new List<int>();

            if (txtSearch.Text.Length < 1) return;

            // se prepara la cadena de busqueda
            string searchValue = txtSearch.Text;

            dataGridPropuestas.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
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

                for (int i = 0; i < dataGridPropuestas.Rows.Count; i++)
                {
                    if (dataGridPropuestas.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                    {
                        busqueda.Add(i);
                    }
                }

                // se selecciona la primera fila encontrada, si existe una
                if (busqueda.Count > 0)
                {
                    iBusqueda = 0;
                    dataGridPropuestas.Rows[busqueda[0]].Selected = true;
                    dataGridPropuestas.FirstDisplayedScrollingRowIndex = dataGridPropuestas.SelectedRows[0].Index;
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

            dataGridPropuestas.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridPropuestas.FirstDisplayedScrollingRowIndex = dataGridPropuestas.SelectedRows[0].Index;
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            // failsafe
            if (busqueda.Count <= 0) return;

            iBusqueda--;
            if (iBusqueda <= 0) iBusqueda = 0;

            dataGridPropuestas.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridPropuestas.FirstDisplayedScrollingRowIndex = dataGridPropuestas.SelectedRows[0].Index;
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
