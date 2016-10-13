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

        private void btnPublicacionesDoct_Click(object sender, EventArgs e)
        {
            padre.AbrirPublicacionesDoctForm();
        }

        private void btnPonenciasDoct_Click(object sender, EventArgs e)
        {
            padre.AbrirPonenciasDoctForm();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            padre.AbrirMatriculaDoctForm();
        }

        private void btnEstudiantesMaes_Click(object sender, EventArgs e)
        {
            padre.AbrirEstudiantesMaesForm();
        }

        private void btnMatriculaMaes_Click(object sender, EventArgs e)
        {
            padre.AbrirMatriculaMaesForm();
        }

        private void btnPublicacionesMaes_Click(object sender, EventArgs e)
        {
            padre.AbrirPublicacionesMaesForm();
        }

        private void btnPonenciasMaes_Click(object sender, EventArgs e)
        {
            padre.AbrirPonenciasMaesForm();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            padre.AbrirPropuestasMaesForm();
        }

        private void btnTesisDoct_Click(object sender, EventArgs e)
        {
            padre.AbrirTesisDoctForm();
        }

        private void btnTesisMaes_Click(object sender, EventArgs e)
        {
            padre.AbrirTesisMaesForm();
        }        
    }
}
