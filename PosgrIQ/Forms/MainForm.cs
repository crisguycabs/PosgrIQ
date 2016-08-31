using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace PosgrIQ
{
    public partial class MainForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia de la ventana ProfesoresForm
        /// </summary>
        public ProfesoresForm profesoresForm = null;

        /// <summary>
        /// Indica si esta abierta, o no, la ventana ProfesoresForm
        /// </summary>
        public bool abiertoProfesoresForm = false;

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
            AbrirHomeForm();
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
