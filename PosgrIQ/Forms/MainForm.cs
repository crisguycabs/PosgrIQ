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
    public partial class MainForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia de la ventana TesisMaesForm
        /// </summary>
        public TesisMaes tesisMaesForm = null;

        /// <summary>
        /// Indica si la ventana TesisMaesForm esta abierta o no
        /// </summary>
        public bool abiertoTesisMaesForm = false;

        /// <summary>
        /// Instancia de la ventana de PropuestaMaesForm
        /// </summary>
        public PropuestaMaesForm propuestaMaesForm = null;

        /// <summary>
        /// Indica si la ventana PropuestaMaesForm esta abierta o no
        /// </summary>
        public bool abiertoPropuestaMaesForm = false;

        /// <summary>
        /// Instancia de la ventana PonenciasMaesForm
        /// </summary>
        public PonenciasMaesForm ponenciasMaesForm = null;

        /// <summary>
        /// Indica si la ventana PonenciasMaesForm esta abierta o no
        /// </summary>
        public bool abiertoPonenciasMaesForm = false;

        /// <summary>
        /// Instancia de la ventana PublicacionesMaesForm
        /// </summary>
        public PublicacionesMaesForm publicacionesMaesForm = null;

        /// <summary>
        /// Indica si la ventana PublicacionesMaesForm esta abierta o no
        /// </summary>
        public bool abiertoPublicacionesMaesForm = false;

        /// <summary>
        /// Instancia de la ventana MatriculaMaesForm
        /// </summary>
        public MatriculaMaesForm matriculaMaesForm = null;

        /// <summary>
        /// Indica si la ventana MatriculaMaesForm esta abierta o no
        /// </summary>
        public bool abiertoMatriculaMaesForm = false;

        /// <summary>
        /// Instancia de la ventana EstudiantesMaesForm
        /// </summary>
        public EstudiantesMaesForm estudiantesMaesForm = null;

        /// <summary>
        /// Indica si la ventana EstudiantesMaesForm esta abierta o no
        /// </summary>
        public bool abiertoEstudiantesMaesForm = false;

        /// <summary>
        /// Instancia de la ventana MatriculaDoctForm
        /// </summary>
        public MatriculaDoctForm matriculaDoctForm = null;

        /// <summary>
        /// Indica si la ventana MatriculaDoctForm esta abierta o no
        /// </summary>
        public bool abiertoMatriculaDoctForm = false;

        /// <summary>
        /// Instancia de la ventana PublicacionesDoctForm
        /// </summary>
        public PonenciasDoctForm ponenciasDoctForm = null;

        /// <summary>
        /// Indica si la ventana PonenciasDoctForm esta abierta o no
        /// </summary>
        public bool abiertoPonenciasDoctForm = false;

        /// <summary>
        /// Instancia de la ventana PublicacionesDoctForm
        /// </summary>
        public PublicacionesDoctForm publicacionesDoctForm = null;

        /// <summary>
        /// Indica si la ventana PublicacionesDoctForm esta abierta o no
        /// </summary>
        public bool abiertoPublicacionesDoctForm = false;

        /// <summary>
        /// Instancia de la ventana ProfesoresForm
        /// </summary>
        public ProfesoresForm profesoresForm = null;

        /// <summary>
        /// Indica si esta abierta, o no, la ventana ProfesoresForm
        /// </summary>
        public bool abiertoProfesoresForm = false;

        /// <summary>
        /// Instancia de la ventana TesisDoctForm
        /// </summary>
        public TesisDoctForm tesisDoctForm = null;

        /// <summary>
        /// Indica si la ventana esta abierta o no
        /// </summary>
        public bool abiertoTesisDoctForm = false;

        /// <summary>
        /// Instancia de la ventana HomeForm
        /// </summary>
        public HomeForm homeForm = null;

        /// <summary>
        /// Indica si esta abierta, o no, la ventana HomeForm
        /// </summary>
        public bool abiertoHomeForm = false;

        /// <summary>
        /// Instancia de la ventana EscuelasForm
        /// </summary>
        public EscuelasForm escuelasForm = null;

        /// <summary>
        /// Indica si esta abierta, o no, la ventana EscuelasForm
        /// </summary>
        public bool abiertoEscuelasForm = false;

        /// <summary>
        /// Instancia de la ventana SemestresForm
        /// </summary>
        public SemestresForm semestresForm = null;

        /// <summary>
        /// Indica si la ventana SemestresForm esta abierta, o no
        /// </summary>
        public bool abiertoSemestresForm = false;

        /// <summary>
        /// Instancia de la ventana ReglamentoForm
        /// </summary>
        public ReglamentosForm reglamentosForm = null;

        /// <summary>
        /// Indica si la ventana ReglamentosForm esta abierta, o no
        /// </summary>
        public bool abiertoReglamentosForm = false;

        /// <summary>
        /// Instancia de la ventana EstudiantesDoctForm
        /// </summary>
        public EstudiantesDoctForm estudiantesDoctForm = null;

        /// <summary>
        /// Indica si la ventana EstudiantesDoctForm esta abierta, o no
        /// </summary>
        public bool abiertoEstudiantesDoctForm=false;

        /// <summary>
        /// Instancia de la ventana PropuestaDoctForm
        /// </summary>
        public PropuestaDoctForm propuestaDoctForm = null;

        /// <summary>
        /// Indica si la ventana PropuestaDoctForm esta abierta, o no
        /// </summary>
        public bool abiertoPropuestaDoctForm = false;

        /// <summary>
        /// Instancia de la ventana ConfiguracionForm
        /// </summary>
        public ConfiguracionForm configuracionForm = null;

        /// <summary>
        /// Indica si la ventana ConfiguracionForm esta abierta, o no
        /// </summary>
        public bool abiertoConfiguracionForm = false;

        /// <summary>
        /// Ruta fisica de la BD
        /// </summary>
        public string sourceBD = "";

        /// <summary>
        /// Ruta fisica de la carpeta de OneDrive
        /// </summary>
        public string sourceONE = "";

        #endregion

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        static public void ErrorOleDb(string tabla, OleDbException ex)
        {
            string errorMessages = "No se puede acceder a la tabla " + tabla + "\r\n\r\n";

            for (int i = 0; i < ex.Errors.Count; i++)
            {
                errorMessages += ex.Errors[i].Message + "\r\n\r\n";                
            }

            MessageBox.Show(errorMessages, "ERROR DE CONEXION", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Se verifica que no exista mas de un archivo .MDB en la carpeta de la base de datos. Es decir, que no hayan copias en conflicto
        /// </summary>
        public void CheckConflicto()
        {
            string folder = System.IO.Path.GetDirectoryName(sourceBD);
            string[] files = System.IO.Directory.GetFiles(folder, "*.mdb", System.IO.SearchOption.AllDirectories);
            if (files.Length > 1)
            {
                // se encontraron más de un archivo MDB
                // se le informa al usuario y se cierra la aplicacion
                MessageBox.Show("Se encontraron " + files.Length + " copias en conflicto de la base de datos en la carpeta " + folder + ".\r\n\r\nElimine las copias en conflicto antes seguir.\r\n\r\nPosgrIQ se cerrará.", "Copias de la BD en conflicto", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            if (!GetSource())
            {
                MessageBox.Show("No se encuentra la ruta de la base de datos.\r\n\r\nLa aplicacion se cerrara.", "Error de archivo de configuracion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            if (!GetOne())
            {
                MessageBox.Show("No se encuentra la ruta de la carpeta de OneDrive.\r\n\r\nLa aplicacion se cerrara.", "Error de archivo de configuracion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            CheckConflicto();
            if (!TestSource())
            {
                MessageBox.Show("La base de datos seleccionada no contiene las tablas necesarias.\r\n\r\nSe debe cambiar la configuracion de la aplicacion.", "Error de archivo de configuracion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AbrirConfiguracionForm(true);
            }
            else AbrirHomeForm();
        }

        /// <summary>
        /// Se toma la ruta del OneDrive desde el archivo sourceone
        /// </summary>
        /// <returns></returns>
        public bool GetOne()
        {
            try
            {
                System.IO.StreamReader sr = new System.IO.StreamReader("sourceone");
                this.sourceONE = sr.ReadLine();
                sr.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Se toma la base de datos desde el archivo source
        /// </summary>
        /// <returns></returns>
        public bool GetSource()
        {
            try
            {
                System.IO.StreamReader sr = new System.IO.StreamReader("source");
                this.sourceBD = sr.ReadLine();
                sr.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Se escribe la base de datos en el archivo source
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public bool SetSource(string source)
        {
            try
            {
                System.IO.StreamWriter sw = new System.IO.StreamWriter("source",false);
                sw.WriteLine(source);
                sw.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Se escribe la ruta de OneDrive en el archivo sourceone
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public bool SetSourceOne(string sourceone)
        {
            try
            {
                System.IO.StreamWriter sw = new System.IO.StreamWriter("sourceone", false);
                sw.WriteLine(sourceone);
                sw.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Se prueba que la base de datos escogida contenga todas las tablas requeridas
        /// </summary>
        /// <returns></returns>
        public bool TestSource()
        {
            try
            {
                var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);
                
                conection.Open();

                // algunas variables
                string query, query2;
                OleDbCommand command;
                OleDbDataAdapter da;

                query2 = "Colegiatura";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Colegiatura";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Conceptos";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Condicion";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Configuracion";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Escuelas";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "EstudiantesDoct";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "EstudiantesMaes";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "MatriculaDoct";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "MatriculaMaes";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "PonenciasDoct";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "PonenciasMaes";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Profesores";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "PropuestaDoct";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "PropuestaMaes";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "PublicacionesDoct";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "PublicacionesMaes";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Reglamentos";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "Semestres";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close(); 
                    return false;
                }

                query2 = "TesisDoct";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close();
                    return false;
                }

                query2 = "TesisMaes";
                query = "SELECT * FROM " + query2 + " ORDER BY codigo ASC";
                try
                {
                    command = new OleDbCommand(query, conection);
                    da = new OleDbDataAdapter(command);
                }
                catch
                {
                    MessageBox.Show("La base de datos no contienen la tabla " + query2 + ".\r\nNo es una base de datos valida", "Error de validacion de BD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    conection.Close();
                    return false;
                }

                conection.Close();
                
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void AbrirConfiguracionForm(bool conf)
        {
            if (!abiertoConfiguracionForm)
            {
                configuracionForm = new ConfiguracionForm();
                configuracionForm.padre = this;
                configuracionForm.MdiParent = this;
                configuracionForm.conf = conf;

                abiertoConfiguracionForm = true;

                configuracionForm.Show();
            }
            else
            {
                configuracionForm.Select();
            }
        }

        public void CerrarConfiguracionForm()
        {
            configuracionForm.Close();
            configuracionForm = null;
            abiertoConfiguracionForm = false;
        }

        public void AbrirReglamentosForm()
        {
            if (!abiertoReglamentosForm)
            {
                reglamentosForm = new ReglamentosForm();
                reglamentosForm.padre = this;
                reglamentosForm.MdiParent = this;

                abiertoReglamentosForm = true;

                reglamentosForm.Show();
            }
            else
            {
                reglamentosForm.Select();
            }
        }

        public void CerrarReglamentosForm()
        {
            reglamentosForm.Close();
            reglamentosForm = null;
            abiertoReglamentosForm = false;
        }

        public void AbrirEstudiantesDoctForm()
        {
            if (!abiertoEstudiantesDoctForm)
            {
                estudiantesDoctForm = new EstudiantesDoctForm();
                estudiantesDoctForm.padre = this;
                estudiantesDoctForm.MdiParent = this;

                abiertoEstudiantesDoctForm = true;

                estudiantesDoctForm.Show();
            }
            else
            {
                estudiantesDoctForm.Select();
            }
        }

        public void CerrarEstudiantesDoctForm()
        {
            estudiantesDoctForm.Close();
            estudiantesDoctForm = null;
            abiertoEstudiantesDoctForm = false;
        }

        public void AbrirEstudiantesMaesForm()
        {
            if (!abiertoEstudiantesMaesForm)
            {
                estudiantesMaesForm = new EstudiantesMaesForm();
                estudiantesMaesForm.padre = this;
                estudiantesMaesForm.MdiParent = this;

                abiertoEstudiantesMaesForm = true;

                estudiantesMaesForm.Show();
            }
            else
            {
                estudiantesMaesForm.Select();
            }
        }

        public void CerrarEstudiantesMaesForm()
        {
            estudiantesMaesForm.Close();
            estudiantesMaesForm = null;
            abiertoEstudiantesMaesForm = false;
        }

        public void AbrirTesisDoctForm()
        {
            if(!abiertoTesisDoctForm)
            {
                tesisDoctForm=new TesisDoctForm();
                tesisDoctForm.padre=this;
                tesisDoctForm.MdiParent=this;

                abiertoTesisDoctForm=true;

                tesisDoctForm.Show();
            }
            else
            {
                tesisDoctForm.Select();
            }
        }

        public void CerrarTesisDoctForm()
        {
            tesisDoctForm.Close();
            tesisDoctForm = null;
            abiertoTesisDoctForm = false;
        }

        public void AbrirTesisMaesForm()
        {
            if (!abiertoTesisMaesForm)
            {
                tesisMaesForm = new TesisMaes();
                tesisMaesForm.padre = this;
                tesisMaesForm.MdiParent = this;

                abiertoTesisMaesForm = true;

                tesisMaesForm.Show();
            }
            else
            {
                tesisMaesForm.Select();
            }
        }

        public void CerrarTesisMaesForm()
        {
            tesisMaesForm.Close();
            tesisMaesForm = null;
            abiertoTesisMaesForm = false;
        }

        public void AbrirPropuestasDoctForm()
        {
            if (!abiertoPropuestaDoctForm)
            {
                propuestaDoctForm = new PropuestaDoctForm();
                propuestaDoctForm.padre = this;
                propuestaDoctForm.MdiParent = this;

                abiertoPropuestaDoctForm = true;

                propuestaDoctForm.Show();
            }
            else
            {
                propuestaDoctForm.Select();
            }
        }

        public void CerrarPropuestasDoctForm()
        {
            propuestaDoctForm.Close();
            propuestaDoctForm = null;
            abiertoPropuestaDoctForm = false;
        }

        public void AbrirPropuestasMaesForm()
        {
            if (!abiertoPropuestaMaesForm)
            {
                propuestaMaesForm = new PropuestaMaesForm();
                propuestaMaesForm.padre = this;
                propuestaMaesForm.MdiParent = this;

                abiertoPropuestaMaesForm = true;

                propuestaMaesForm.Show();
            }
            else
            {
                propuestaDoctForm.Select();
            }
        }

        public void CerrarPropuestasMaesForm()
        {
            propuestaMaesForm.Close();
            propuestaMaesForm = null;
            abiertoPropuestaMaesForm = false;
        }

        public void AbrirEscuelasForm()
        {
            if (!abiertoEscuelasForm)
            {
                escuelasForm = new EscuelasForm();
                escuelasForm.padre = this;
                escuelasForm.MdiParent = this;

                abiertoEscuelasForm = true;

                escuelasForm.Show();
            }
            else
            {
                escuelasForm.Select();
            }
        }

        public void CerrarEscuelasForm()
        {
            escuelasForm.Close();
            escuelasForm = null;
            abiertoEscuelasForm = false;
        }

        public void AbrirMatriculaDoctForm()
        {
            if (!abiertoMatriculaDoctForm)
            {
                matriculaDoctForm = new MatriculaDoctForm();
                matriculaDoctForm.padre = this;
                matriculaDoctForm.MdiParent = this;

                abiertoMatriculaDoctForm = true;

                matriculaDoctForm.Show();
            }
            else
            {
                matriculaDoctForm.Select();
            }
        }

        public void CerrarMatriculaDoctForm()
        {
            matriculaDoctForm.Close();
            matriculaDoctForm = null;
            abiertoMatriculaDoctForm = false;
        }

        public void AbrirMatriculaMaesForm()
        {
            if (!abiertoMatriculaMaesForm)
            {
                matriculaMaesForm = new MatriculaMaesForm();
                matriculaMaesForm.padre = this;
                matriculaMaesForm.MdiParent = this;

                abiertoMatriculaMaesForm = true;

                matriculaMaesForm.Show();
            }
            else
            {
                matriculaMaesForm.Select();
            }
        }

        public void CerrarMatriculaMaesForm()
        {
            matriculaMaesForm.Close();
            matriculaMaesForm = null;
            abiertoMatriculaMaesForm = false;
        }

        public void AbrirPublicacionesDoctForm()
        {
            if (!abiertoPublicacionesDoctForm)
            {
                publicacionesDoctForm = new PublicacionesDoctForm();
                publicacionesDoctForm.padre = this;
                publicacionesDoctForm.MdiParent = this;

                abiertoPublicacionesDoctForm = true;

                publicacionesDoctForm.Show();
            }
            else
            {
                publicacionesDoctForm.Select();
            }
        }

        public void CerrarPublicacionesDoctForm()
        {
            publicacionesDoctForm.Close();
            publicacionesDoctForm = null;
            abiertoPublicacionesDoctForm = false;
        }

        public void AbrirPublicacionesMaesForm()
        {
            if (!abiertoPublicacionesMaesForm)
            {
                publicacionesMaesForm = new PublicacionesMaesForm();
                publicacionesMaesForm.padre = this;
                publicacionesMaesForm.MdiParent = this;

                abiertoPublicacionesMaesForm = true;

                publicacionesMaesForm.Show();
            }
            else
            {
                publicacionesMaesForm.Select();
            }
        }

        public void CerrarPublicacionesMaesForm()
        {
            publicacionesMaesForm.Close();
            publicacionesMaesForm = null;
            abiertoPublicacionesMaesForm = false;
        }

        public void AbrirPonenciasDoctForm()
        {
            if (!abiertoPonenciasDoctForm)
            {
                ponenciasDoctForm = new PonenciasDoctForm();
                ponenciasDoctForm.padre = this;
                ponenciasDoctForm.MdiParent = this;

                abiertoPonenciasDoctForm = true;

                ponenciasDoctForm.Show();
            }
            else
            {
                ponenciasDoctForm.Select();
            }
        }

        public void CerrarPonenciasDoctForm()
        {
            ponenciasDoctForm.Close();
            ponenciasDoctForm = null;
            abiertoPonenciasDoctForm = false;
        }

        public void AbrirPonenciasMaesForm()
        {
            if (!abiertoPonenciasMaesForm)
            {
                ponenciasMaesForm = new PonenciasMaesForm();
                ponenciasMaesForm.padre = this;
                ponenciasMaesForm.MdiParent = this;

                abiertoPonenciasMaesForm = true;

                ponenciasMaesForm.Show();
            }
            else
            {
                ponenciasDoctForm.Select();
            }
        }

        public void CerrarPonenciasMaesForm()
        {
            ponenciasMaesForm.Close();
            ponenciasMaesForm = null;
            abiertoPonenciasMaesForm = false;
        }

        public void AbrirSemestresForm()
        {
            if (!abiertoSemestresForm)
            {
                semestresForm = new SemestresForm();
                semestresForm.padre = this;
                semestresForm.MdiParent = this;

                abiertoSemestresForm = true;

                semestresForm.Show();
            }
            else
            {
                semestresForm.Select();
            }
        }

        public void CerrarSemestresForm()
        {
            semestresForm.Close();
            semestresForm = null;
            abiertoSemestresForm = false;
        }

        public void AbrirHomeForm()
        {
            if (!abiertoHomeForm)
            {
                homeForm = new HomeForm();
                homeForm.padre = this;
                homeForm.MdiParent = this;

                abiertoHomeForm = true;

                homeForm.Show();
            }
            else
            {
                homeForm.Select();
            }
        }

        public void CerrarHomeForm()
        {
            homeForm.Close();
            homeForm = null;
            abiertoHomeForm = false;
        }

        public void AbrirProfesoresForm()
        {
            if (!abiertoProfesoresForm)
            {
                profesoresForm = new ProfesoresForm();
                profesoresForm.padre = this;
                profesoresForm.MdiParent = this;

                abiertoProfesoresForm = true;

                profesoresForm.Show();
            }
            else
            {
                profesoresForm.Select();
            }
        }

        public void CerrarProfesoresForm()
        {
            profesoresForm.Close();
            profesoresForm = null;
            abiertoProfesoresForm = false;
        }

        public static int ReescalarAltoDataGridView(DataGridView control)
        {
            return control.Rows.Cast<DataGridViewRow>().Sum(x => x.Height) + (control.RowHeadersVisible ? 10 : 0) + 14;                
        }

        public static void ReescalarDataGridView(ref DataGridView control)
        {
            if (MainForm.ReescalarAltoDataGridView(control) > 200) control.Height = 200;
            else control.Height = MainForm.ReescalarAltoDataGridView(control);
        }

        /// <summary>
        /// Se convierte una fecha en texto dd/mm/aaaa a un formato DateTime
        /// </summary>
        /// <param name="texto"></param>
        /// <returns></returns>
        public static DateTime Texto2Fecha(string texto)
        {
            string[] texto2 = texto.Split('/');

            if (texto2.Length < 3) return DateTime.Now;
            return new DateTime(Convert.ToInt32(texto2[2]), Convert.ToInt32(texto2[1]), Convert.ToInt32(texto2[0]));
        }

        /// <summary>
        /// Convierte un formato fecha a texto dd/mm/aaaa
        /// </summary>
        /// <param name="fecha"></param>
        /// <returns></returns>
        public static String Fecha2Texto(DateTime fecha)
        {
            string mes = fecha.Month.ToString();
            string dia = fecha.Day.ToString();

            if (mes.Length < 2) mes = "0" + mes;
            if (dia.Length < 2) dia = "0" + dia;

            return (dia + "/" + mes + "/" + fecha.Year.ToString());
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void cerrarTodoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (abiertoEstudiantesDoctForm) CerrarEstudiantesDoctForm();
            if (abiertoEstudiantesMaesForm) CerrarEstudiantesMaesForm();

            if (abiertoMatriculaDoctForm) CerrarMatriculaDoctForm();
            if (abiertoMatriculaMaesForm) CerrarMatriculaMaesForm();

            if (abiertoPonenciasDoctForm) CerrarPonenciasDoctForm();
            if (abiertoPonenciasMaesForm) CerrarPonenciasMaesForm();

            if (abiertoPropuestaDoctForm) CerrarPropuestasDoctForm();
            if (abiertoPropuestaMaesForm) CerrarPropuestasMaesForm();

            if (abiertoPublicacionesDoctForm) CerrarPublicacionesDoctForm();
            if (abiertoPublicacionesDoctForm) CerrarPublicacionesMaesForm();

            if (abiertoTesisDoctForm) CerrarPublicacionesDoctForm();
            if (abiertoTesisMaesForm) CerrarPublicacionesMaesForm();

            if (abiertoEscuelasForm) CerrarEscuelasForm();
            if (abiertoProfesoresForm) CerrarProfesoresForm();
            if (abiertoProfesoresForm) CerrarProfesoresForm();
            if (abiertoReglamentosForm) CerrarReglamentosForm();
            if (abiertoSemestresForm) CerrarSemestresForm();

            if (abiertoConfiguracionForm) CerrarConfiguracionForm();
        }

        private void menuAcerca_Click(object sender, EventArgs e)
        {
            AboutForm acercaForm = new AboutForm();
            acercaForm.ShowDialog();
        }

        private void menuConfiguracion_Click(object sender, EventArgs e)
        {
            AbrirConfiguracionForm(false);
        }

        private void menuEstudiantesDoct_Click(object sender, EventArgs e)
        {
            AbrirEstudiantesDoctForm();
        }

        private void menuMatriculaDoct_Click(object sender, EventArgs e)
        {
            AbrirMatriculaDoctForm();
        }

        private void menuPonenciasDoct_Click(object sender, EventArgs e)
        {
            AbrirPonenciasDoctForm();
        }

        private void menuPropuestasDoct_Click(object sender, EventArgs e)
        {
            AbrirPropuestasDoctForm();
        }

        private void menuPublicacionesDoct_Click(object sender, EventArgs e)
        {
            AbrirPublicacionesDoctForm();
        }

        private void menuTesisDoct_Click(object sender, EventArgs e)
        {
            AbrirTesisDoctForm();
        }

        private void menuEstudiantesMaes_Click(object sender, EventArgs e)
        {
            AbrirEstudiantesMaesForm();
        }

        private void menuMatriculaMaes_Click(object sender, EventArgs e)
        {
            AbrirMatriculaMaesForm();
        }

        private void menuPonenciasMaes_Click(object sender, EventArgs e)
        {
            AbrirPonenciasMaesForm();
        }

        private void menuPropuestasMaes_Click(object sender, EventArgs e)
        {
            AbrirPropuestasMaesForm();
        }

        private void menuPublicacionesMaes_Click(object sender, EventArgs e)
        {
            AbrirPublicacionesMaesForm();
        }

        private void menuTesisMaes_Click(object sender, EventArgs e)
        {
            AbrirPublicacionesMaesForm();
        }

        private void escuelasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AbrirEscuelasForm();
        }

        private void profesoresToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AbrirProfesoresForm();
        }

        private void reglamentosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AbrirReglamentosForm();
        }

        private void semestresToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AbrirSemestresForm();
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los estudiantes de maestria
        /// </summary>
        public void InformeEstudiantesMaes()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dt = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\EstudiantesMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dt = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesMaes WHERE condicion=1 ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dt);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los reglamentos
            DataTable dtReglamentos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Reglamentos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtReglamentos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Reglamentos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de las condiciones
            DataTable dtCondicion = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Condicion ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtCondicion);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Condicion", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            /*  CODIGO
                NOMBRE 
                CORREO
                CEDULA
                CIUDAD
                CONDICION
                NIVEL
                DIRECTOR
                CODIRECTOR 1
                CODIRECTOR 2
                REGLAMENTO
                TEMA
                FECHA
                CONCEPTO
                OBSERVACIONES
            */

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowpos = 5;
                colpos = (3 * i) + 1;

                ws.Columns[colpos - 1].Width = 5 * 300;
                ws.Columns[colpos].Width = 15 * 300;
                ws.Columns[colpos + 1].Width = 42 * 300;

                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(i + 1) + "/" + dt.Rows.Count.ToString();
                ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIGO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][0].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "NOMBRE";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][1].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CORREO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CEDULA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][3].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CIUDAD";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][4].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONDICION";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtCondicion.Select("codigo=" + Convert.ToString(dt.Rows[i][5]))[0][1]);                
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "NIVEL";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][6].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "DIRECTOR";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe director?

                if (dt.Rows[i][7].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][7]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value="";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIRECTOR 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe codirector1?
                if (dt.Rows[i][8].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][8]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIRECTOR 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe codirector2?
                if (dt.Rows[i][9].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][9]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "REGLAMENTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtReglamentos.Select("codigo=" + Convert.ToString(dt.Rows[i][10]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "TEMA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][11].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][12].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dt.Rows[i][13]))[0][1]);
                //ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][13].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "OBSERVACIONES";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][15].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;
                ws.Cells[rowpos, colpos].Value = "";
                rowpos++;
                ws.Cells[rowpos, colpos].Value = "";
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\EstudiantesMaes.pdf");
            
            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());

            // se crea el archivo mas grande de estudiantes de maestria
            // se borra el archivo creado
            System.IO.File.Delete(final);

            // se lee la plantilla
            inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes2.xlsx";

            // se duplica el archivo
            System.IO.File.Copy(inicial, final);

            // se vuelve a escribir el archivo final
            ef = ExcelFile.Load(final);
            ws = ef.Worksheets[0];

            /*  CODIGO
                NOMBRE 
                CORREO
                CEDULA
                CIUDAD
                CONDICION
                NIVEL
                DIRECTOR
                CODIRECTOR 1
                CODIRECTOR 2
                REGLAMENTO
                TEMA
                FECHA
                CONCEPTO
                OBSERVACIONES
            */

            // se escriben los encabezados

            colpos = 1;
            rowpos = 5;

            ws.Cells[rowpos, colpos].Value = "CODIGO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "NOMBRE";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CORREO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CEDULA";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CIUDAD";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CONDICION";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "NIVEL";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "DIRECTOR";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CODIRECTOR 1";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CODIRECTOR 2";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "REGLAMENTO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "TEMA";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "FECHA";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CONCEPTO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "OBSERVACIONES";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowpos = 6+i;
                colpos = 1;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][0].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][1].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][3].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][4].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = Convert.ToString(dtCondicion.Select("codigo=" + Convert.ToString(dt.Rows[i][5]))[0][1]); 
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][6].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                // existe director?
                if (dt.Rows[i][7].ToString() != "") ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][7]))[0][1]);
                else ws.Cells[rowpos, colpos].Value = "";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                colpos++;

                // existe codirector1?
                if (dt.Rows[i][8].ToString() != "") ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][8]))[0][1]);
                else ws.Cells[rowpos, colpos].Value = "";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                colpos++;

                // existe codirector2?
                if (dt.Rows[i][9].ToString() != "") ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][9]))[0][1]);
                else ws.Cells[rowpos, colpos].Value = "";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                colpos++;

                ws.Cells[rowpos, colpos].Value = Convert.ToString(dtReglamentos.Select("codigo=" + Convert.ToString(dt.Rows[i][10]))[0][1]);
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][11].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][12].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dt.Rows[i][13]))[0][1]);
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][15].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los estudiantes de doctorado
        /// </summary>
        public void InformeEstudiantesDoct()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dt = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\EstudiantesDoct.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de doctorado
            dt = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesDoct WHERE condicion=1 ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dt);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los reglamentos
            DataTable dtReglamentos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Reglamentos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtReglamentos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Reglamentos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de la condicion
            DataTable dtCondicion = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Condicion ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtCondicion);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Condicion", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            /*  CODIGO
                NOMBRE 
                CORREO
                CEDULA
                CIUDAD
                CONDICION
                NIVEL
                DIRECTOR
                CODIRECTOR 1
                CODIRECTOR 2
                REGLAMENTO
                TEMA
                FECHA
                CONCEPTO
                OBSERVACIONES
            */

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowpos = 5;
                colpos = (3 * i) + 1;

                ws.Columns[colpos - 1].Width = 5 * 300;
                ws.Columns[colpos].Width = 15 * 300;
                ws.Columns[colpos + 1].Width = 42 * 300;

                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(i + 1) + "/" + dt.Rows.Count.ToString();
                ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIGO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][0].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "NOMBRE";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][1].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CORREO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CEDULA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][3].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CIUDAD";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][4].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONDICION";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][5].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "NIVEL";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][6].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "DIRECTOR";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe director?
                if (dt.Rows[i][7].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][7]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIRECTOR 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe codirector1?
                if (dt.Rows[i][8].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][8]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIRECTOR 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe codirector2?
                if (dt.Rows[i][9].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][9]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "REGLAMENTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtReglamentos.Select("codigo=" + Convert.ToString(dt.Rows[i][10]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "TEMA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][11].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][12].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dt.Rows[i][13]))[0][1]);
                //ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][13].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "SOLICITO QUALIFY";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][15].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "APROBO QUALIFY";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][16].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "OBSERVACIONES";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][17].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;                               
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\EstudiantesDoct.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());

            // se crea el archivo mas grande de estudiantes de maestria
            // se borra el archivo creado
            System.IO.File.Delete(final);

            // se lee la plantilla
            inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes2.xlsx";

            // se duplica el archivo
            System.IO.File.Copy(inicial, final);

            // se vuelve a escribir el archivo final
            ef = ExcelFile.Load(final);
            ws = ef.Worksheets[0];

            // se escriben los encabezados

            colpos = 1;
            rowpos = 5;

            ws.Cells[rowpos, colpos].Value = "CODIGO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "NOMBRE";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CORREO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CEDULA";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CIUDAD";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CONDICION";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "NIVEL";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "DIRECTOR";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CODIRECTOR 1";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CODIRECTOR 2";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "REGLAMENTO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "TEMA";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "FECHA";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "CONCEPTO";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "SOLICITO QUALIFY";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "APROBO QUALIFY";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            colpos++;

            ws.Cells[rowpos, colpos].Value = "OBSERVACIONES";
            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowpos = 6 + i;
                colpos = 1;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][0].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][1].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][3].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][4].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = Convert.ToString(dtCondicion.Select("codigo=" + Convert.ToString(dt.Rows[i][5]))[0][1]);
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][6].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                // existe director?
                if (dt.Rows[i][7].ToString() != "") ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][7]))[0][1]);
                else ws.Cells[rowpos, colpos].Value = "";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                colpos++;

                // existe codirector1?
                if (dt.Rows[i][8].ToString() != "") ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][8]))[0][1]);
                else ws.Cells[rowpos, colpos].Value = "";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                colpos++;

                // existe codirector2?
                if (dt.Rows[i][9].ToString() != "") ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dt.Rows[i][9]))[0][1]);
                else ws.Cells[rowpos, colpos].Value = "";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                colpos++;

                ws.Cells[rowpos, colpos].Value = Convert.ToString(dtReglamentos.Select("codigo=" + Convert.ToString(dt.Rows[i][10]))[0][1]);
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][11].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][12].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dt.Rows[i][13]))[0][1]);
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][15].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][16].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

                colpos++;

                ws.Cells[rowpos, colpos].Value = dt.Rows[i][17].ToString();
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;

            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de las propuestas de maestria
        /// </summary>
        public void InformePropuestaMaes()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\PropuestasMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtPropuestas = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM PropuestaMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtPropuestas);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla PropuestaMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }       

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            for (int i = 0; i < dtPropuestas.Rows.Count; i++)
            {
                rowpos = 5;
                colpos = (3 * i) + 1;

                ws.Columns[colpos - 1].Width = 5 * 300;
                ws.Columns[colpos].Width = 18 * 300;
                ws.Columns[colpos + 1].Width = 42 * 300;

                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(i + 1) + "/" + dtEstudiantes.Rows.Count.ToString();
                ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][1]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "TITULO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][4]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][5]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA ENTREGA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][6].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][7].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][7]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][8].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][8]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA CORRECCIONES";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][11].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][12].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][12]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][13].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][13]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA SUSTENTACION";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][16].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][17].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][17]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;                
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\PropuestasMaes.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de las propuestas de doctorado
        /// </summary>
        public void InformePropuestaDoct()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\PropuestasDoct.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtPropuestas = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM PropuestaDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtPropuestas);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla PropuestaDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];
            
            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            for (int i = 0; i < dtPropuestas.Rows.Count; i++)
            {
                rowpos = 5;
                colpos = (3 * i) + 1;

                ws.Columns[colpos - 1].Width = 5 * 300;
                ws.Columns[colpos].Width = 18 * 300;
                ws.Columns[colpos + 1].Width = 42 * 300;

                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(i + 1) + "/" + dtEstudiantes.Rows.Count.ToString();
                ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][1]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "TITULO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][4]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][5]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 3";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][6]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 4";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe calificador4
                if (dtPropuestas.Rows[i][7].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][7]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA ENTREGA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][8].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][9].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][9]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][10].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][10]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 3";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][11].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][11]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 4";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][12].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][12]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA CORRECCIONES";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][17].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][18].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][18]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][19].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][19]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 3";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][20].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][20]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 4";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][21].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][21]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA SUSTENTACION";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtPropuestas.Rows[i][26].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtPropuestas.Rows[i][27].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtPropuestas.Rows[i][27]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\PropuestasDoct.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de las tesis de maestria
        /// </summary>
        public void InformeTesisMaes()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\TesisMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtTesis = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM TesisMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtTesis);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla TesisMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            for (int i = 0; i < dtTesis.Rows.Count; i++)
            {
                rowpos = 5;
                colpos = (3 * i) + 1;

                ws.Columns[colpos - 1].Width = 5 * 300;
                ws.Columns[colpos].Width = 18 * 300;
                ws.Columns[colpos + 1].Width = 42 * 300;

                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(i + 1) + "/" + dtEstudiantes.Rows.Count.ToString();
                ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][1]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "TITULO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][4]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][5]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA ENTREGA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][6].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][7].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][7]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][8].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][8]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA CORRECCIONES";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][11].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][12].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][12]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][13].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][13]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA SUSTENTACION";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][16].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][17].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][17]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\TesisMaes.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de las tesis de doctorado
        /// </summary>
        public void InformeTesisDoct()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\TesisDoct.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtTesis = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM TesisDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtTesis);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla TesisDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se verifica si no hay tesis a guardar
            if (dtTesis.Rows.Count == 0) return;

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            for (int i = 0; i < dtTesis.Rows.Count; i++)
            {
                rowpos = 5;
                colpos = (3 * i) + 1;

                ws.Columns[colpos - 1].Width = 5 * 300;
                ws.Columns[colpos].Width = 18 * 300;
                ws.Columns[colpos + 1].Width = 42 * 300;

                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(i + 1) + "/" + dtEstudiantes.Rows.Count.ToString();
                ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][1]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "TITULO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][2].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Style.WrapText = true;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][4]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][5]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 3";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][6]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 4";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][7]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CALIFICADOR 5";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtProfesores.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][8]))[0][1]);
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA ENTREGA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][9].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][10].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][10]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][11].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][11]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 3";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][12].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][12]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 4";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][13].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][13]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 5";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][14].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][14]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA CORRECCIONES";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][20].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][21].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][21]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][22].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][22]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 3";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][23].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][23]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 4";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][24].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][24]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO CAL 5";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][25].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][25]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "FECHA SUSTENTACION";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dtTesis.Rows[i][31].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CONCEPTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                // existe concepto1?
                if (dtTesis.Rows[i][32].ToString() != "") ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtConceptos.Select("codigo=" + Convert.ToString(dtTesis.Rows[i][32]))[0][1]);
                else ws.Cells[rowpos, colpos + 1].Value = "";
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\TesisDoct.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los calificadores de propuestas de maestria
        /// </summary>
        public void InformeCalificadoresPropMaes()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\CalifPropuestasMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateCalificadores.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtPropuestas = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM PropuestaMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtPropuestas);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla PropuestaMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            int nprof = 0;
            
            for (int i = 0; i < dtProfesores.Rows.Count; i++)
            {
                // se extraen todos las propuestas de maestria en el que el calificador aparezca
                DataRow[] dr = dtPropuestas.Select("calificador1=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador2=" + Convert.ToString(dtProfesores.Rows[i][0]));

                try 
                {
                    if(dr.Length>0)
                    {
                        // el profesor es calificador de alguna propuesta, por tanto se escribe en el Excel

                        rowpos = 4;
                        colpos = (nprof * 4) + 1;
                        nprof++;

                        ws.Columns[colpos - 1].Width = 2 * 300;
                        ws.Columns[colpos].Width = 30 * 300;
                        ws.Columns[colpos + 1].Width = 83 * 300;
                        ws.Columns[colpos + 2].Width = 2 * 300;

                        ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Rows[i][1]).ToUpper();
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

                        rowpos++;

                        ws.Cells[rowpos, colpos - 1].Value = "#";
                        ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos + 1].Value = "TITULO";
                        ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                        for(int j=0;j<dr.Length;j++)
                        {
                            rowpos++;

                            ws.Cells[rowpos, colpos - 1].Value = (j+1).ToString();
                            ws.Cells[rowpos, colpos - 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            ws.Cells[rowpos, colpos].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + dr[j][1])[0][1]);
                            ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                            ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dr[j][2]).Substring(0,70)+"...";
                            ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.NormalWeight;
                        }
                    }
                }
                catch
                { }                
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;
            ws.PrintOptions.RightMargin = 1.5;
                        
            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\CalifPropuestasMaes.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los calificadores de propuestas de doctorado
        /// </summary>
        public void InformeCalificadoresPropDoct()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\CalifPropuestasDoct.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateCalificadores.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtPropuestas = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM PropuestaDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtPropuestas);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla PropuestaDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            int nprof = 0;

            for (int i = 0; i < dtProfesores.Rows.Count; i++)
            {
                // se extraen todos las propuestas de maestria en el que el calificador aparezca
                DataRow[] dr = dtPropuestas.Select("calificador1=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador2=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador3=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador4=" + Convert.ToString(dtProfesores.Rows[i][0]));

                try
                {
                    if (dr.Length > 0)
                    {
                        // el profesor es calificador de alguna propuesta, por tanto se escribe en el Excel

                        rowpos = 4;
                        colpos = (nprof * 4) + 1;
                        nprof++;

                        ws.Columns[colpos - 1].Width = 2 * 300;
                        ws.Columns[colpos].Width = 30 * 300;
                        ws.Columns[colpos + 1].Width = 83 * 300;
                        ws.Columns[colpos + 2].Width = 2 * 300;

                        ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Rows[i][1]).ToUpper();
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

                        rowpos++;

                        ws.Cells[rowpos, colpos - 1].Value = "#";
                        ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos + 1].Value = "TITULO";
                        ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                        for (int j = 0; j < dr.Length; j++)
                        {
                            rowpos++;

                            ws.Cells[rowpos, colpos - 1].Value = (j + 1).ToString();
                            ws.Cells[rowpos, colpos - 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            ws.Cells[rowpos, colpos].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + dr[j][1])[0][1]);
                            ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                            ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dr[j][2]).Substring(0, 70) + "...";
                            ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.NormalWeight;
                        }
                    }
                }
                catch
                { }
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;
            ws.PrintOptions.RightMargin = 1.5;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\CalifPropuestasDoct.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los calificadores de propuestas de maestria
        /// </summary>
        public void InformeCalificadoresTesisMaes()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\CalifTesisMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateCalificadores.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtTesis = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM TesisMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtTesis);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla TesisMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            int nprof = 0;

            for (int i = 0; i < dtProfesores.Rows.Count; i++)
            {
                // se extraen todos las propuestas de maestria en el que el calificador aparezca
                DataRow[] dr = dtTesis.Select("calificador1=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador2=" + Convert.ToString(dtProfesores.Rows[i][0]));

                try
                {
                    if (dr.Length > 0)
                    {
                        // el profesor es calificador de alguna propuesta, por tanto se escribe en el Excel

                        rowpos = 4;
                        colpos = (nprof * 4) + 1;
                        nprof++;

                        ws.Columns[colpos - 1].Width = 2 * 300;
                        ws.Columns[colpos].Width = 30 * 300;
                        ws.Columns[colpos + 1].Width = 83 * 300;
                        ws.Columns[colpos + 2].Width = 2 * 300;

                        ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Rows[i][1]).ToUpper();
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

                        rowpos++;

                        ws.Cells[rowpos, colpos - 1].Value = "#";
                        ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos + 1].Value = "TITULO";
                        ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                        for (int j = 0; j < dr.Length; j++)
                        {
                            rowpos++;

                            ws.Cells[rowpos, colpos - 1].Value = (j + 1).ToString();
                            ws.Cells[rowpos, colpos - 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            ws.Cells[rowpos, colpos].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + dr[j][1])[0][1]);
                            ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                            ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dr[j][2]).Substring(0, 70) + "...";
                            ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.NormalWeight;
                        }
                    }
                }
                catch
                { }
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;
            ws.PrintOptions.RightMargin = 1.5;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\CalifTesisMaes.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los calificadores de propuestas de maestria
        /// </summary>
        public void InformeCalificadoresTesisDoct()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\CalifTesisDoct.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateCalificadores.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtTesis = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM TesisDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtTesis);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla TesisDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // no hay registros de tesis
            if (dtTesis.Rows.Count == 0) return;

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            int nprof = 0;

            for (int i = 0; i < dtProfesores.Rows.Count; i++)
            {
                // se extraen todos las propuestas de maestria en el que el calificador aparezca
                DataRow[] dr = dtTesis.Select("calificador1=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador2=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador3=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador4=" + Convert.ToString(dtProfesores.Rows[i][0]) + " OR calificador5=" + Convert.ToString(dtProfesores.Rows[i][0]));

                try
                {
                    if (dr.Length > 0)
                    {
                        // el profesor es calificador de alguna propuesta, por tanto se escribe en el Excel

                        rowpos = 4;
                        colpos = (nprof * 4) + 1;
                        nprof++;

                        ws.Columns[colpos - 1].Width = 2 * 300;
                        ws.Columns[colpos].Width = 30 * 300;
                        ws.Columns[colpos + 1].Width = 83 * 300;
                        ws.Columns[colpos + 2].Width = 2 * 300;

                        ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Rows[i][1]).ToUpper();
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

                        rowpos++;

                        ws.Cells[rowpos, colpos - 1].Value = "#";
                        ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos + 1].Value = "TITULO";
                        ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                        for (int j = 0; j < dr.Length; j++)
                        {
                            rowpos++;

                            ws.Cells[rowpos, colpos - 1].Value = (j + 1).ToString();
                            ws.Cells[rowpos, colpos - 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            ws.Cells[rowpos, colpos].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + dr[j][1])[0][1]);
                            ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                            ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dr[j][2]).Substring(0, 70) + "...";
                            ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.NormalWeight;
                        }
                    }
                }
                catch
                { }
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;
            ws.PrintOptions.RightMargin = 1.5;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\CalifTesisDoct.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los directores de maestria
        /// </summary>
        public void InformeDirectorMaes()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\DirectorMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateCalificadores.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //// se carga la informacion de las tesis de maestria
            //DataTable dtTesis = new DataTable();
            //try
            //{
            //    conection.Open();

            //    // se pide la informacion de los estudiantes de maestria
            //    query = "SELECT * FROM TesisMaes ORDER BY codigo ASC";
            //    command = new OleDbCommand(query, conection);

            //    da = new OleDbDataAdapter(command);
            //    da.Fill(dtTesis);

            //    conection.Close();
            //}
            //catch
            //{
            //    MessageBox.Show("No se encuentra la tabla TesisMaes", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            int nprof = 0;

            for (int i = 0; i < dtProfesores.Rows.Count; i++)
            {
                // se extraen todos los estudiantes de maestria en el que el profesor aparezca como director
                DataRow[] dr = dtEstudiantes.Select("director=" + Convert.ToString(dtProfesores.Rows[i][0]));

                try
                {
                    if (dr.Length > 0)
                    {
                        // el profesor es director de alguna estudiante de maestria, por tanto se escribe en el Excel

                        rowpos = 4;
                        colpos = (nprof * 4) + 1;
                        nprof++;

                        ws.Columns[colpos - 1].Width = 2 * 300;
                        ws.Columns[colpos].Width = 30 * 300;
                        ws.Columns[colpos + 1].Width = 83 * 300;
                        ws.Columns[colpos + 2].Width = 2 * 300;

                        ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Rows[i][1]).ToUpper();
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

                        rowpos++;

                        ws.Cells[rowpos, colpos - 1].Value = "#";
                        ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos + 1].Value = "TITULO";
                        ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                        for (int j = 0; j < dr.Length; j++)
                        {
                            rowpos++;

                            ws.Cells[rowpos, colpos - 1].Value = (j + 1).ToString();
                            ws.Cells[rowpos, colpos - 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            ws.Cells[rowpos, colpos].Value = Convert.ToString(dr[j][1]);
                            ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                            ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtEstudiantes.Select("codigo="+ Convert.ToString(dr[j][0]))[0][11]);
                            //ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtTesis.Select("estudiante=" + Convert.ToString(dr[j][0]))[0][2]);
                            //Convert.ToString(dr[j][2]).Substring(0, 70) + "...";
                            ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.NormalWeight;
                        }
                    }
                }
                catch
                { }
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;
            ws.PrintOptions.RightMargin = 1.5;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\DirectorMaes.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

        /// <summary>
        /// Genera el informe en PDF y XLSX de los calificadores de propuestas de maestria
        /// </summary>
        public void InformeDirectorDoct()
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dtEstudiantes = new DataTable();
            OleDbDataAdapter da;

            // se lee desde el archivo oneSource la ruta a donde se deben guardar el archivo excel
            System.IO.StreamReader sr;

            try
            {
                sr = new System.IO.StreamReader("sourceone");
            }
            catch
            {
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string leido = @sr.ReadLine();
            string final = leido + "\\DirectorDoct.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateCalificadores.xlsx";

            // se verifica que existan los archivos
            //bool existeInicial = false;
            //existeInicial = System.IO.File.Exists(inicial);
            //bool existeFin = false;
            //existeFin = System.IO.File.Exists(final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (System.IO.File.Exists(final)) System.IO.File.Delete(final);
            }
            catch
            {
                MessageBox.Show("El archivo " + final + "esta siendo utilizado y no se puede modificar.\r\n\r\nPor favor cierre el archivo e intentelo de nuevo.", "Error de lectura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // se duplica el archivo
                System.IO.File.Copy(inicial, final);
            }
            catch
            {
                MessageBox.Show("No se puede crear el nuevo archivo excel en la ruta " + final, "Error de escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dtEstudiantes = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtEstudiantes);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los profesores
            DataTable dtProfesores = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtProfesores);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Profesores", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtConceptos = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtConceptos);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla Conceptos", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los conceptos
            DataTable dtTesis = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM TesisDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dtTesis);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla TesisDoct", "Error en la Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int rowpos;
            int colpos;

            int nprof = 0;

            for (int i = 0; i < dtProfesores.Rows.Count; i++)
            {
                // se extraen todos las propuestas de maestria en el que el calificador aparezca
                DataRow[] dr = dtEstudiantes.Select("director=" + Convert.ToString(dtProfesores.Rows[i][0]));

                try
                {
                    if (dr.Length > 0)
                    {
                        // el profesor es director de alguna tesis, por tanto se escribe en el Excel

                        rowpos = 4;
                        colpos = (nprof * 4) + 1;
                        nprof++;

                        ws.Columns[colpos - 1].Width = 2 * 300;
                        ws.Columns[colpos].Width = 30 * 300;
                        ws.Columns[colpos + 1].Width = 83 * 300;
                        ws.Columns[colpos + 2].Width = 2 * 300;

                        ws.Cells[rowpos, colpos].Value = Convert.ToString(dtProfesores.Rows[i][1]).ToUpper();
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;

                        rowpos++;

                        ws.Cells[rowpos, colpos - 1].Value = "#";
                        ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos].Value = "ESTUDIANTE";
                        ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                        ws.Cells[rowpos, colpos + 1].Value = "TITULO";
                        ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.BoldWeight;

                        for (int j = 0; j < dr.Length; j++)
                        {
                            rowpos++;

                            ws.Cells[rowpos, colpos - 1].Value = (j + 1).ToString();
                            ws.Cells[rowpos, colpos - 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos - 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            ws.Cells[rowpos, colpos].Value = Convert.ToString(dr[j][1]);
                            ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.NormalWeight;
                            ws.Cells[rowpos, colpos + 1].Value = Convert.ToString(dtEstudiantes.Select("codigo=" + Convert.ToString(dr[j][0]))[0][11]);
                            //Convert.ToString(dr[j][2]).Substring(0, 70) + "...";
                            ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                            ws.Cells[rowpos, colpos + 1].Style.Font.Weight = ExcelFont.NormalWeight;
                        }
                    }
                }
                catch
                { }
            }

            //ef.Save(final);

            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitToPage = false;
            ws.PrintOptions.TopMargin = 0.2;
            ws.PrintOptions.BottomMargin = 0.2;
            ws.PrintOptions.LeftMargin = 0.2;
            ws.PrintOptions.RightMargin = 1.5;

            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal = System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\DirectorDoct.pdf");

            DateTime end = DateTime.Now;
            //MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }
    }
}
