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
    public partial class SplashForm : Form
    {
        #region

        /// <summary>
        /// temporizador para el splash screen
        /// </summary>
        int timeleft;
        
        #endregion

        public SplashForm()
        {
            InitializeComponent();
        }

        private void SplashForm_Load(object sender, EventArgs e)
        {
            timeleft = 1;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (timeleft > 0)
            {
                timeleft = timeleft - 1;

            }
            else
            {
                timer1.Stop();
                new MainForm().Show();
                this.Hide();
            }
        }
    }
}
