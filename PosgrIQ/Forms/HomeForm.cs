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
    public partial class HomeForm : Form
    {
        #region variables de clase

        public MainForm padre;

        #endregion

        public HomeForm()
        {
            InitializeComponent();
        }

        private void btnProfesores_Click(object sender, EventArgs e)
        {
            padre.AbrirProfesoresForm();
        }

        private void btnEscuelas_Click(object sender, EventArgs e)
        {
            padre.AbrirEscuelasForm();
        }

        private void btnSemestres_Click(object sender, EventArgs e)
        {
            padre.AbrirSemestresForm();
        }

        private void btnReglamentos_Click(object sender, EventArgs e)
        {
            padre.AbrirReglamentosForm();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            padre.AbrirEstudiantesDoctForm();
        }

        private void btnPropuestasDoct_Click(object sender, EventArgs e)
        {
            padre.AbrirPropuestasDoctForm();
        }

        private void btnConfiguracion_Click(object sender, EventArgs e)
        {
            padre.AbrirConfiguracionForm(false);
        }        
    }
}
