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
    public partial class AddReglamentosForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia al MainForm padre
        /// </summary>
        public MainForm padre;

        /// <summary>
        /// Valores posibles: '1' para agregar, '0' para modificar, '2' para agregar invocado desde otra ventana
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

        public AddReglamentosForm()
        {
            InitializeComponent();
        }

        private void AddReglamentosForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            // se lee desde la BD la cantidad de escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Reglamentos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dt = new DataTable();
                da.Fill(dt);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                padre.ShowWaiting("Espere un momento mientras PosgrIQ procesa...");
                System.Threading.Thread.Sleep(1000);
                padre.CloseWaiting();

                switch (modo)
                {
                    case true: // se agrega un nuevo reglamento

                        // se genera el nuevo codigo para el nuevo reglamento
                        codigo = dt.Rows.Count + 1;
                        txtNombre.Text = "";

                        this.Text = "AGREGAR NUEVO REGLAMENTO";

                        break;

                    case false: // se modifica un reglamento

                        // se escriben en los controles la informacion del reglamento a modificar
                        txtNombre.Text = Convert.ToString(dt.Select("codigo=" + codigo.ToString())[0][1]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR REGLAMENTO";

                        break;
                    default:
                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Reglamentos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // se realizan algunas comprobaciones de seguridad

            // nombre vacio
            if (string.IsNullOrWhiteSpace(txtNombre.Text))
            {
                MessageBox.Show("No se ha introducido el nombre del reglamento", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // existe un reglamento con ese codigo. Solo revisar en modo insercion
            if (modo)
            {
                DataRow[] busqueda = dt.Select("codigo=" + codigo.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe un reglamento con el codigo " + codigo.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }


            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query;
            OleDbCommand command;

            switch (modo)
            {
                case true:

                    // se agrega el reglamento
                    try
                    {
                        // se prepara la cadena SQL
                        query = "INSERT INTO Reglamentos VALUES(" + codigo.ToString() + ",'" + txtNombre.Text + "')";
                        
                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        padre.ShowWaiting("Espere un momento mientras PosgrIQ procesa...");
                        System.Threading.Thread.Sleep(1000);
                        padre.CloseWaiting();

                        this.txtNombre.Text = "";
                        codigo++;

                        if (modo) this.DialogResult = DialogResult.OK;
                        else
                        {
                            try
                            {
                                padre.reglamentosForm.ReglamentosForm_Load(sender, e);
                            }
                            catch { }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Reglamentos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                case false:

                    // se modifica la escuela
                    try
                    {
                        // se prepara la cadena SQL
                        query = "UPDATE Reglamentos SET codigo=" + codigo.ToString() + ", nombre='" + txtNombre.Text + "' WHERE codigo=" + codigo.ToString();

                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        padre.ShowWaiting("Espere un momento mientras PosgrIQ procesa...");
                        System.Threading.Thread.Sleep(1000);
                        padre.CloseWaiting();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Reglamentos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;

                default:
                    break;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
