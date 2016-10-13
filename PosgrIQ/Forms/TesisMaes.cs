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
    public partial class TesisMaes : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public TesisMaes()
        {
            InitializeComponent();
        }

        private void TesisMaes_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtTesisMaes, dtProfesores, dtConceptos, dtEstudiantesMaes;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM TesisMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtTesisMaes = new DataTable();
                da.Fill(dtTesisMaes);

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtEstudiantesMaes = new DataTable();
                da.Fill(dtEstudiantesMaes);

                // se pide la informacion de los profesores
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                // se pide la informacion de los conceptos
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtConceptos = new DataTable();
                da.Fill(dtConceptos);

                // se crea una tabla que reemplaza el codigo de colegiatura y de escuela por el nombre
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Estudiante", typeof(string));
                dt.Columns.Add("Titulo", typeof(string));
                dt.Columns.Add("Fecha Entrega", typeof(string));
                dt.Columns.Add("Calificador 1", typeof(string));
                dt.Columns.Add("Calificador 2", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 2", typeof(string));
                dt.Columns.Add("Fecha Entrega Correcciones", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 2", typeof(string));
                dt.Columns.Add("Fecha Sustentacion", typeof(string));
                dt.Columns.Add("Concepto Final", typeof(string));

                // se llena el nuevo datatable
                for (int i = 0; i < dtTesisMaes.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtTesisMaes.Rows[i][0];

                    // nombre del estudiante
                    DataRow[] seleccionado = dtEstudiantesMaes.Select("codigo=" + dtTesisMaes.Rows[i][1]);
                    fila[1] = Convert.ToString(seleccionado[0][1]);

                    // titulo de la propuesta
                    fila[2] = dtTesisMaes.Rows[i][2];

                    // fecha de entrega
                    fila[3] = Convert.ToString(dtTesisMaes.Rows[i][6].ToString());

                    // calificador 1
                    fila[4] = dtProfesores.Select("codigo=" + dtTesisMaes.Rows[i][4].ToString())[0][1];

                    // calificador 2
                    fila[5] = dtProfesores.Select("codigo=" + dtTesisMaes.Rows[i][5].ToString())[0][1];

                    // concepto 1 calificador 1
                    fila[6] = dtConceptos.Select("codigo=" + dtTesisMaes.Rows[i][7].ToString())[0][1];

                    // concepto 1 calificador 2
                    fila[7] = dtConceptos.Select("codigo=" + dtTesisMaes.Rows[i][8].ToString())[0][1];

                    // fecha de entrega de correcciones
                    fila[8] = dtTesisMaes.Rows[i][11];

                    // concepto 2 calificador 1
                    fila[9] = dtConceptos.Select("codigo=" + dtTesisMaes.Rows[i][12].ToString())[0][1];

                    // concepto 2 calificador 2
                    fila[10] = dtConceptos.Select("codigo=" + dtTesisMaes.Rows[i][13].ToString())[0][1];

                    // fecha sustentacion
                    fila[11] = dtTesisMaes.Rows[i][16];

                    // concepto sustentacion
                    fila[12] = dtTesisMaes.Rows[i][17];

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridTesis.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridTesis);

                dataGridTesis.Sort(dataGridTesis.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Tesis Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            int codigo = Convert.ToInt32(dataGridTesis.SelectedRows[0].Cells[0].Value);
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dtTesisMaes;
                OleDbDataAdapter da;

                // se pide la informacion de la propuesta de doct
                query = "SELECT * FROM TesisMaes WHERE codigo=" + codigo.ToString() + " ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtTesisMaes = new DataTable();
                da.Fill(dtTesisMaes);

                conection.Close();

                try
                {
                    switch (cmbVer.SelectedIndex)
                    {
                        case 0: // propuesta
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisMaes.Rows[0][3]));
                            break;
                        case 1: // concepto 1 calificador 1
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisMaes.Rows[0][10]));
                            break;
                        case 2: // concepto 1 calificador 2
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisMaes.Rows[0][11]));
                            break;
                        case 3: // concepto 2 calificador 1
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisMaes.Rows[0][21]));
                            break;
                        case 4: // concepto 2 calificador 2
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisMaes.Rows[0][22]));
                            break;
                        case 5: // sustentacion
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisMaes.Rows[0][33]));
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
                MessageBox.Show("No se puede acceder a la base de datos, tabla Tesis Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddTesisMaesForm agregar = new AddTesisMaesForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.TesisMaes_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridTesis.Rows.Count < 1) return;

            AddTesisMaesForm modificar = new AddTesisMaesForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridTesis.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.TesisMaes_Load(sender, e);
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarTesisMaesForm();
        }
    }
}
