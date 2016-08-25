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
        /// Valores posibles: 'true' para agregar, 'false' para modificar
        /// </summary>
        public bool modo;

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
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                string query = "SELECT * FROM Escuelas";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                
                dtEscuelas = new DataTable();
                da.Fill(dtEscuelas);

                conection.Close();

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
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                string query = "SELECT * FROM Colegiatura";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtColegiatura = new DataTable();
                da.Fill(dtColegiatura);

                conection.Close();

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
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Profesores";
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
                    case true: // se agrega un nuevo profesor

                        numCod.Value = 0;
                        txtNombre.Text = "";
                        txtCorreo.Text = "";
                        txtTelefono.Text = "";
                        txtUniversidad.Text = "";
                        
                        this.Text = "AGREGAR NUEVO PROFESOR";

                        break;

                    case false: // se modifica una escuela

                        // se escriben en los controles la informacion de la escuela seleccionada

                        DataRow[] seleccionado = dtProfesores.Select("codigo=" + codigo.ToString());

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

            // nombre vacio
            if (string.IsNullOrWhiteSpace(txtNombre.Text))
            {
                MessageBox.Show("No se ha introducido el nombre del profesor", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // existe un profesor con ese codigo. Solo revisar en modo insercion
            if (modo)
            {
                DataRow[] busqueda = dtProfesores.Select("codigo=" + numCod.Value.ToString());
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
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            string query;
            OleDbCommand command;

            switch (modo)
            {
                case true:

                    // se agrega el profesor
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "INSERT INTO Profesores VALUES(" + numCod.Value.ToString() + ",'" + txtNombre.Text + "',"+(cmbColegiatura.SelectedIndex+1).ToString()+","+(cmbEscuela.SelectedIndex+1).ToString()+",'"+ txtCorreo.Text +"','"+ txtTelefono.Text +"','"+ txtUniversidad.Text + "')";
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                case false:

                    // se modifica el profesor
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "UPDATE Profesores SET codigo=" + numCod.Value.ToString() + ", nombre='" + txtNombre.Text + "', colegiatura=" + (cmbColegiatura.SelectedIndex + 1).ToString() + ", escuela=" + (cmbEscuela.SelectedIndex + 1).ToString() + ", correo='" + txtCorreo.Text + "', telefono='" + txtTelefono.Text + "', universidad='" + txtUniversidad.Text + "' WHERE codigo=" + codigo.ToString();
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                default:
                    break;
            }            
        }
    }
}
