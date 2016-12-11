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
    public partial class AddProfesoresForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia al MainForm padre
        /// </summary>
        public MainForm padre;

        /// <summary>
        /// Valores posibles: '1' para agregar, '0' para modificar, '2' para agregar invocado desde otra ventana
        /// </summary>
        public int modo;

        /// <summary>
        /// Codigo del profesor a modificar
        /// </summary>
        public int codigo;

        /// <summary>
        /// Guarda la informacion de la tabla Profesores antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtProfesores;

        /// <summary>
        /// Guarda la informacion de la tabla Colegiatura antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtColegiatura;

        /// <summary>
        /// Guarda la informacion de la tabla Escuelas antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtEscuelas;

        #endregion

        public AddProfesoresForm()
        {
            InitializeComponent();
        }

        private void LlenarEscuelas()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                string query = "SELECT * FROM Escuelas ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtEscuelas = new DataTable();
                da.Fill(dtEscuelas);

                conection.Close();

                cmbEscuela.Items.Clear();

                foreach (DataRow row in dtEscuelas.Rows)
                {
                    cmbEscuela.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Escuelas", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LlenarColegiatura()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                string query = "SELECT * FROM Colegiatura ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtColegiatura = new DataTable();
                da.Fill(dtColegiatura);

                conection.Close();

                cmbColegiatura.Items.Clear();

                foreach (DataRow row in dtColegiatura.Rows)
                {
                    cmbColegiatura.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Colegiatura", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddProfesoresForm_Load(object sender, EventArgs e)
        {
            // se lee desde la BD la cantidad de Profesores, Colegiatura y Escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                conection.Close();

                // se llenan las colegiaturas y escuelas
                LlenarEscuelas();
                LlenarColegiatura();

                switch (modo)
                {
                    case 1:
                    case 2: // se agrega un nuevo profesor

                        numCod.Value = dtProfesores.Rows.Count+1;
                        txtNombre.Text = "";
                        txtCorreo.Text = "";
                        txtTelefono.Text = "";
                        txtUniversidad.Text = "";

                        this.Text = "AGREGAR NUEVO PROFESOR";

                        break;

                    case 0: // se modifica una escuela

                        // se escriben en los controles la informacion de la escuela seleccionada

                        DataRow[] seleccionado = dtProfesores.Select("codigo=" + codigo.ToString());
                        DataRow[] listaEscuelas = dtEscuelas.Select("codigo=" + seleccionado[0][3]);
                        numCod.Value = codigo;
                        txtNombre.Text = Convert.ToString(seleccionado[0][1]);
                        cmbColegiatura.SelectedIndex = Convert.ToInt32(seleccionado[0][2]) - 1;
                        cmbEscuela.SelectedIndex = Convert.ToInt32(seleccionado[0][3]) - 1;
                        txtCorreo.Text = Convert.ToString(seleccionado[0][4]);
                        txtTelefono.Text = Convert.ToString(seleccionado[0][5]);
                        txtUniversidad.Text = Convert.ToString(seleccionado[0][6]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR PROFESOR";

                        break;
                    default:
                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // se realizan algunas comprobaciones de seguridad
            DataRow[] busqueda;

            // nombre vacio
            if (string.IsNullOrWhiteSpace(txtNombre.Text))
            {
                MessageBox.Show("No se ha introducido el nombre del profesor", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // existe un profesor con ese codigo. Solo revisar en modo insercion
            if (modo>=1)
            {
                busqueda = dtProfesores.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe un profesor con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no se ha seleccionado colegiatura
            if (cmbColegiatura.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado un grado de colegiatura", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // no se ha seleccionado escuela
            if (cmbEscuela.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado una escuela asociada", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // no se ha ingresado un correo
            if (string.IsNullOrWhiteSpace(txtCorreo.Text))
            {
                MessageBox.Show("No se ha introducido el correo electronico", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query;
            OleDbCommand command;

            switch (modo)
            {
                case 1:
                case 2:

                    // se agrega el profesor
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        busqueda = dtEscuelas.Select("escuela='" + cmbEscuela.Items[cmbEscuela.SelectedIndex] + "'");
                        query = "INSERT INTO Profesores VALUES(" + numCod.Value.ToString() + ",'" + txtNombre.Text + "'," + (cmbColegiatura.SelectedIndex + 1).ToString() + "," + busqueda[0][0] + ",'" + txtCorreo.Text + "','" + txtTelefono.Text + "','" + txtUniversidad.Text + "')";
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        if (modo == 2) this.DialogResult = DialogResult.OK;
                        else
                        {
                            try
                            {
                                padre.profesoresForm.ProfesoresForm_Load(sender, e);
                            }
                            catch { }
                        }

                        numCod.Value = 0;
                        txtNombre.Text = "";
                        cmbColegiatura.SelectedIndex = -1;
                        cmbEscuela.SelectedIndex = -1;
                        txtCorreo.Text = "";
                        txtTelefono.Text = "";
                        txtUniversidad.Text = "";

                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                case 0:

                    // se modifica el profesor
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        busqueda = dtEscuelas.Select("escuela='" + cmbEscuela.Items[cmbEscuela.SelectedIndex] + "'");
                        query = "UPDATE Profesores SET codigo=" + numCod.Value.ToString() + ", nombre='" + txtNombre.Text + "', colegiatura=" + (cmbColegiatura.SelectedIndex + 1).ToString() + ", escuela=" + busqueda[0][0] + ", correo='" + txtCorreo.Text + "', telefono='" + txtTelefono.Text + "', universidad='" + txtUniversidad.Text + "' WHERE codigo=" + codigo.ToString();
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                default:
                    break;
            }
        }

        private void btnAddEscuela_Click(object sender, EventArgs e)
        {
            AddEscuelasForm agregar = new AddEscuelasForm();
            agregar.padre = this.padre;
            agregar.modo = 2;

            if (agregar.ShowDialog() == DialogResult.OK)
            {
                this.LlenarEscuelas();
                this.cmbEscuela.SelectedIndex = this.cmbEscuela.Items.Count - 1;
            }
        }
    }
}
