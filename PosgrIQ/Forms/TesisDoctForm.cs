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
    public partial class TesisDoctForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public TesisDoctForm()
        {
            InitializeComponent();
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            int codigo = Convert.ToInt32(dataGridTesis.SelectedRows[0].Cells[0].Value);
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dtTesisDoct;
                OleDbDataAdapter da;

                // se pide la informacion de la propuesta de doct
                query = "SELECT * FROM TesisDoct WHERE codigo=" + codigo.ToString() + " ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtTesisDoct = new DataTable();
                da.Fill(dtTesisDoct);

                conection.Close();

                try
                {
                    switch (cmbVer.SelectedIndex)
                    {
                        case 0: // propuesta
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][3]));
                            break;
                        case 1: // concepto 1 calificador 1
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][10]));
                            break;
                        case 2: // concepto 1 calificador 2
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][11]));
                            break;
                        case 3: // concepto 1 calificador 3
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][12]));
                            break;
                        case 4: // concepto 1 calificador 4
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][13]));
                            break;
                        case 5: // concepto 1 calificador 5
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][14]));
                            break;
                        case 6: // concepto 2 calificador 1
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][21]));
                            break;
                        case 7: // concepto 2 calificador 2
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][22]));
                            break;
                        case 8: // concepto 2 calificador 3
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][23]));
                            break;
                        case 9: // concepto 2 calificador 4
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][24]));
                            break;
                        case 10: // concepto 2 calificador 5
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][25]));
                            break;
                        case 11: // sustentacion
                            System.Diagnostics.Process.Start(Convert.ToString(dtTesisDoct.Rows[0][33]));
                            break;
                    }
                }
                catch
                {
                    MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla TesisDoct", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarTesisDoctForm();
        }
    }
}
