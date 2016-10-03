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
    public partial class AddMatriculaDoctForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia al MainForm padre
        /// </summary>
        public MainForm padre;

        /// <summary>
        /// Valores posibles: 'true' para agregar, 'false' para modificar
        /// </summary>
        public bool modo;

        /// <summary>
        /// Codigo de la propuesta a modificar
        /// </summary>
        public int codigo;

        /// <summary>
        /// Guarda la informacion de la tabla Estudiantes antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtEstudiantes;

        /// <summary>
        /// Guarda la informacion de la tabla MatriculaDoct antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtMatriculaDoct;

        /// <summary>
        /// Guarda la informacion de la tabla Semestres antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtSemestres;

        #endregion

        public AddMatriculaDoctForm()
        {
            InitializeComponent();
        }
    }
}
