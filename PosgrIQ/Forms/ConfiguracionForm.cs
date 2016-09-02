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
    public partial class ConfiguracionForm : Form
    {
        #region variables de clase

        public MainForm padre;

        /// <summary>
        /// indica si la ventana fue llamada desde el inicio de la aplicación debido a un error de configuracion
        /// </summary>
        public bool conf = false;

        #endregion

        public ConfiguracionForm()
        {
            InitializeComponent();
        }

        private void chkCambios_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCambios.Checked)
            {
                // si se ha invocado la configuracion desde el inicio de la aplicacion, no es necesario hacer el doble chequeo
                if (!conf)
                {
                    if (MessageBox.Show("Esta seguro que quiere editar esta informacion?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Hand) == DialogResult.Yes)
                    {
                        if (MessageBox.Show("Puede ser una muy mala idea, en serio esta seguro?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            MessageBox.Show("Aun esta a tiempo de evitar daños.\n\r\n\rVuelva a desmarcar la casilla.");
                            txtRutaBD.Enabled = btnRutaBD.Enabled = txtClave.Enabled = txtCorreo.Enabled = txtRutaOne.Enabled = btnRutaOne.Enabled = chkCambios.Checked;
                        }
                    }
                }
                else txtRutaBD.Enabled = btnRutaBD.Enabled = txtClave.Enabled = txtCorreo.Enabled = txtRutaOne.Enabled = btnRutaOne.Enabled = chkCambios.Checked;
            }
            else txtRutaBD.Enabled = btnRutaBD.Enabled = txtClave.Enabled = txtCorreo.Enabled = txtRutaOne.Enabled = btnRutaOne.Enabled = chkCambios.Checked;
        }

        private void ConfiguracionForm_Load(object sender, EventArgs e)
        {
            if (conf)
            {
                // se ha invocado la configuracion desde el inicio de la aplicacion
                txtRutaBD.Enabled = true;
                btnRutaBD.Enabled = true;

                txtRutaBD.Text = padre.sourceBD;
            }
            else
            {
                // se carga la info de configuracion desde la base de datos y el archivo source, funcionamiento normal
                var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                try
                {
                    conection.Open();

                    // algunas variables
                    string query;
                    OleDbCommand command;
                    DataTable dt;
                    OleDbDataAdapter da;

                    // se pide la informacion de los profesores
                    query = "SELECT * FROM Configuracion";
                    command = new OleDbCommand(query, conection);

                    da = new OleDbDataAdapter(command);
                    dt = new DataTable();
                    da.Fill(dt);

                    conection.Close();

                    txtCorreo.Text = dt.Rows[0][0].ToString();
                    txtClave.Text = dt.Rows[0][1].ToString();
                    txtRutaOne.Text = dt.Rows[0][2].ToString();

                    txtRutaBD.Text = padre.sourceBD;
                }
                catch
                {
                    MessageBox.Show("Error al acceder a la base de datos.\r\n\r\nNo se puede guardar la informacion de configuracion.", "Error de acceso BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            // se guarda la configuracion en la base de datos y en el archivo source

            if (txtRutaBD.Enabled)
            {
                // siempre se guarda la ruta de la base de datos
                padre.sourceBD = txtRutaBD.Text;
                padre.SetSource(txtRutaBD.Text);
                padre.AbrirHomeForm();
            }

            if (!conf & chkCambios.Checked)
            {
                // se guarda el resto de la informacion de configuracion si la ventana ha sido invocada desde el HomeForm
                try
                {
                    var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);

                    conection.Open();

                    // algunas variables
                    string query;
                    OleDbCommand command;
                    
                    // se prepara la cadena SQL
                    query = "UPDATE Configuracion SET correo='" + txtCorreo.Text + "', clave='" + txtClave.Text + "', rutaone='" + txtRutaOne.Text + "'";
                    command = new OleDbCommand(query, conection);

                    command.ExecuteNonQuery();

                    conection.Close();
                }
                catch
                {
                    MessageBox.Show("Error al acceder a la base de datos.\r\n\r\nNo se puede guardar la informacion de configuracion.", "Error de acceso BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }            

            padre.CerrarConfiguracionForm();
        }

        private void btnRutaBD_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos MDB (.mdb)|*.mdb";
            open.FilterIndex = 0;
            open.FileName = padre.sourceBD;

            if (open.ShowDialog() == DialogResult.OK)
            {
                padre.sourceBD = open.FileName;
                if (padre.TestSource())
                {
                    MessageBox.Show("La base de datos seleccionada es valida", "Validacion correcta de BD",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    txtRutaBD.Text = open.FileName;
                }
            }
        }

        private void btnRutaOne_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.Description = "Seleccione la carpeta de OneDrive en su PC";
            folder.ShowNewFolderButton = false;

            if (folder.ShowDialog() == DialogResult.OK)
            {
                txtRutaOne.Text = folder.SelectedPath;
            }
        }
    }
}
