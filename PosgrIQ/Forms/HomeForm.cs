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
            string leido = @sr.ReadLine();
            string final = leido + "\\EstudiantesMaes.xlsx";

            // se extrae la ruta inicial
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string inicial = System.IO.Path.GetDirectoryName(path) + "\\templates\\templateEstudiantesMaes.xlsx";

            // se verifica que existan los archivos
            bool existeInicial = false;
            existeInicial = System.IO.File.Exists(inicial);
            bool existeFin = false;
            existeFin = System.IO.File.Exists(final);

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
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
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
                rowpos = 9;
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
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][7].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIRECTOR 1";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][8].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "CODIRECTOR 2";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][9].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "REGLAMENTO";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][10].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "TEMA";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos+1].Value = dt.Rows[i][11].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos+1].Style.WrapText = true;

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
                ws.Cells[rowpos, colpos + 1].Value = dt.Rows[i][13].ToString();
                ws.Cells[rowpos, colpos + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                rowpos++;

                ws.Cells[rowpos, colpos].Value = "OSERVACIONES";
                ws.Cells[rowpos, colpos].Style.Font.Weight = ExcelFont.BoldWeight;
                ws.Cells[rowpos, colpos].Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                ws.Cells[rowpos, colpos+1].Value = dt.Rows[i][15].ToString();
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
            
            ef.Save(final);
            /*
            
            ws.PrintOptions.FitToPage = true;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;            
            */
            //ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "C" + inipos.ToString()));

            string pathFinal=System.IO.Path.GetDirectoryName(final);
            ExcelFile.Load(final).Save(pathFinal + "\\EstudiantesMaes.pdf");
            DateTime end = DateTime.Now;

            MessageBox.Show(((end - start).Milliseconds + (1000 * (end - start).Seconds)).ToString());
        }

    }
}