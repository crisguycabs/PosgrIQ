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
    public partial class AddEscuelasForm : Form
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
        /// Codigo de la escuela a modificar
        /// </summary>
        public int codigo;

        /// <summary>
        /// Guarda la informacion de la tabla Escuelas antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dt;

        #endregion

        public AddEscuelasForm()
        {
            InitializeComponent();
        }

        private void AddEscuelasForm_Load(object sender, EventArgs e)
        {
            // se lee desde la BD la cantidad de escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Escuelas";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dt = new DataTable();
                da.Fill(dt);

                conection.Close();

                switch (modo)
                {
                    case true: // se agrega una nueva escuela

                        // se genera el nuevo codigo para la nueva escuela
                        numCod.Value = dt.Rows.Count + 1;
                        txtNombre.Text = "";

                        this.Text = "Agregar Nueva Escuela";

                        break;

                    case false: // se modifica una escuela

                        // se escriben en los controles la informacion de la escuela seleccionada
                        numCod.Value = codigo;
                        txtNombre.Text = Convert.ToString(dt.Select("codigo=" + codigo.ToString())[0][1]);

                        btnAdd.Text = "Modificar";
                        this.Text = "Modificar Escuela";

                        break;
                    default:
                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // se realizan algunas comprobaciones de seguridad

            // nombre vacio
            if (string.IsNullOrWhiteSpace(txtNombre.Text))
            {
                MessageBox.Show("No se ha introducido el nombre de la escuela", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // existe una escuela con ese codigo. Solo revisar en modo insercion
            if (modo)
            {
                DataRow[] busqueda = dt.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una escuela con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            string query;
            OleDbCommand command;

            switch (modo)
            {
                case true:

                    // se agrega la escuela
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "INSERT INTO Escuelas VALUES(" + numCod.Value.ToString() + ",'" + txtNombre.Text + "')";
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                case false:

                    // se modifica la escuela
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "UPDATE Escuelas SET codigo=" + numCod.Value.ToString() + ", escuela='" + txtNombre.Text + "' WHERE codigo=" + codigo.ToString();
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                default:
                    break;
            }

            this.DialogResult = DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
