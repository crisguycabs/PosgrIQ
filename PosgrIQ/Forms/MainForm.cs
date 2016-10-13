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

        #endregion

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            if (!GetSource())
            {
                MessageBox.Show("No se encuentra la ruta de la base de datos.\r\n\r\nLa aplicacion se cerrara.", "Error de archivo de configuracion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            if (!TestSource())
            {
                MessageBox.Show("No se encuentra la base de datos en la ruta indicada.\r\n\r\nSe debe cambiar la configuracion de la aplicacion.", "Error de archivo de configuracion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AbrirConfiguracionForm(true);
            }
            else AbrirHomeForm();
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
                tesisDoctForm.Select();
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
    }
}
