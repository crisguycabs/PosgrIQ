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
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;

namespace PosgrIQ
{
    public partial class EscuelasForm : Form
    {
        #region variables de clase

        public MainForm padre;

        #endregion

        public EscuelasForm()
        {
            InitializeComponent();
        }
        
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarEscuelasForm();
        }

        public void EscuelasForm_Load(object sender, EventArgs e)
        {
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
                query = "SELECT * FROM Escuelas ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dt = new DataTable();
                da.Fill(dt);

                
                // se enlaza el datatable con el datagrid
                dataGridEscuelas.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridEscuelas);

                dataGridEscuelas.Sort(dataGridEscuelas.Columns[0], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Escuelas", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddEscuelasForm agregar = new AddEscuelasForm();
            agregar.padre = this.padre;
            agregar.modo = 1;

            if (agregar.ShowDialog() == DialogResult.OK) this.EscuelasForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridEscuelas.Rows.Count < 1) return;
            
            AddEscuelasForm modificar = new AddEscuelasForm();
            modificar.padre = this.padre;
            modificar.modo = 0;
            modificar.codigo = Convert.ToInt32(dataGridEscuelas.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.EscuelasForm_Load(sender, e);
        }

        private void dataGridEscuelas_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
    }
}
