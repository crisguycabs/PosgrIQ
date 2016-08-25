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
    }
}
