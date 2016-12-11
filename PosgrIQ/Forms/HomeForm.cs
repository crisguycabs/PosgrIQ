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

        private void HomeForm_Paint(object sender, PaintEventArgs e)
        {
            // ControlPaint.DrawBorder(e.Graphics, e.ClipRectangle, Color.DarkRed, 2, ButtonBorderStyle.Solid, Color.DarkRed, 2, ButtonBorderStyle.Solid, Color.DarkRed, 2, ButtonBorderStyle.Solid, Color.DarkRed, 2, ButtonBorderStyle.Solid);
        }

        private void HomeForm_Load(object sender, EventArgs e)
        {
            label1.BackColor = label3.BackColor = Color.DarkRed;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);

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
                MessageBox.Show("No se puede encontrar el archivo sourceone.\r\n\r\nPor favor verifique la ubicacion del archivo en la ventana de Configuracion e intentelo de nuevo.","Error de lectura",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            // se extrae la ruta de la carpeta final y se prepara el nombre del archivo a copiar
            string final = sr.ReadLine() + "\\EstudiantesMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            bool existeInicial = false;
            existeInicial = System.IO.File.Exists(inicial);
            bool existeFin = false;
            existeFin = System.IO.File.Exists(final);

            System.Diagnostics.Process.Start("explorer.exe", final);

            // si el archivo final existe entonces se debe borrar
            try
            {
                if (existeFin) System.IO.File.Delete(final);
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
                MessageBox.Show("No se puede crear el nuevo archivo excel","Error de escritura",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            // se carga la informacion de los estudiantes de maestria
            dt = new DataTable();
            try
            {
                conection.Open();

                // se pide la informacion de los estudiantes de maestria
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                da.Fill(dt);

                conection.Close();
            }
            catch
            {
                MessageBox.Show("No se encuentra la tabla EstudiantesMaes","Error en la Base de Datos",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            // se abre el archivo excel
            DateTime start = DateTime.Now;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = ExcelFile.Load(final);
            ExcelWorksheet ws = ef.Worksheets[0];

            // se empieza a llenar el archivo de excel
            int inipos = 9;

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
                ws.Cells[inipos, 1].Value = "CODIGO";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][0].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "NOMBRE";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][1].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "CORREO";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][2].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "CEDULA";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][3].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "CIUDAD";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][4].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "CONDICION";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][5].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "DIRECTOR";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][6].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "CODIRECTOR 1";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][7].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "CODIRECTOR 2";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][8].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "REGLAMENTO";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][9].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "TEMA";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][10].ToString();
                ws.Cells[inipos, 2].Style.WrapText = true;

                inipos++;

                ws.Cells[inipos, 1].Value = "FECHA";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][11].ToString();

                inipos++;

                ws.Cells[inipos, 1].Value = "CONCEPTO";
                ws.Cells[inipos, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[inipos, 2].Value = dt.Rows[i][12].ToString();

                inipos++;
                ws.Cells[inipos, 1].Value = "";
                inipos++;
                ws.Cells[inipos, 1].Value = "";
            }

            ef.Save(final);

            ws.PrintOptions.Portrait=true;
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            

            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            
            string pathFinal=System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\EstudiantesMaes.pdf");
            DateTime end = DateTime.Now;

            MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

    }
}