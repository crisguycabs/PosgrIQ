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
        /// Valores posibles: '1' para agregar, '0' para modificar, '2' para agregar invocado desde otra ventana
        /// </summary>
        public int modo;

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
            padre.CheckConflicto();

            // se lee desde la BD la cantidad de escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Escuelas ORDER BY codigo ASC";

                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dt = new DataTable();
                da.Fill(dt);

                switch (modo)
                {
                    case 1:
                    case 2:// se agrega una nueva escuela

                        // se genera el nuevo codigo para la nueva escuela
                        numCod.Value = dt.Rows.Count + 1;
                        txtNombre.Text = "";

                        this.Text = "AGREGAR NUEVA ESCUELA";

                        break;

                    case 0: // se modifica una escuela

                        // se escriben en los controles la informacion de la escuela seleccionada
                        numCod.Value = codigo;
                        txtNombre.Text = Convert.ToString(dt.Select("codigo=" + codigo.ToString())[0][1]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR ESCUELA";

                        break;
                    default:
                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Escuelas", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (modo>=1)
            {
                DataRow[] busqueda = dt.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una escuela con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }


            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query;
            OleDbCommand command;

            switch (modo)
            {
                case 1:
                case 2:

                    // se agrega la escuela
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "INSERT INTO Escuelas VALUES(" + numCod.Value.ToString() + ",'" + txtNombre.Text + "')";
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.txtNombre.Text = "";

                        if (modo == 2) this.DialogResult = DialogResult.OK;
                        else
                        {
                            try { padre.escuelasForm.EscuelasForm_Load(sender, e); }
                            catch { }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Escuelas", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                case 0:

                    // se modifica la escuela
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "UPDATE Escuelas SET codigo=" + numCod.Value.ToString() + ", escuela='" + txtNombre.Text + "' WHERE codigo=" + codigo.ToString();
                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Escuelas", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                default:
                    break;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
