namespace PosgrIQ
{
    partial class HomeForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HomeForm));
            this.btnProfesores = new System.Windows.Forms.Button();
            this.btnEscuelas = new System.Windows.Forms.Button();
            this.btnSemestres = new System.Windows.Forms.Button();
            this.btnReglamentos = new System.Windows.Forms.Button();
            this.btnEstudiantesDoct = new System.Windows.Forms.Button();
            this.btnPropuestasDoct = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnProfesores
            // 
            this.btnProfesores.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnProfesores.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnProfesores.Location = new System.Drawing.Point(50, 77);
            this.btnProfesores.Name = "btnProfesores";
            this.btnProfesores.Size = new System.Drawing.Size(89, 39);
            this.btnProfesores.TabIndex = 0;
            this.btnProfesores.Text = "Profesores";
            this.btnProfesores.UseVisualStyleBackColor = false;
            this.btnProfesores.Click += new System.EventHandler(this.btnProfesores_Click);
            // 
            // btnEscuelas
            // 
            this.btnEscuelas.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnEscuelas.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEscuelas.Location = new System.Drawing.Point(50, 122);
            this.btnEscuelas.Name = "btnEscuelas";
            this.btnEscuelas.Size = new System.Drawing.Size(89, 39);
            this.btnEscuelas.TabIndex = 1;
            this.btnEscuelas.Text = "Escuelas";
            this.btnEscuelas.UseVisualStyleBackColor = false;
            this.btnEscuelas.Click += new System.EventHandler(this.btnEscuelas_Click);
            // 
            // btnSemestres
            // 
            this.btnSemestres.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnSemestres.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSemestres.Location = new System.Drawing.Point(50, 167);
            this.btnSemestres.Name = "btnSemestres";
            this.btnSemestres.Size = new System.Drawing.Size(89, 39);
            this.btnSemestres.TabIndex = 2;
            this.btnSemestres.Text = "Semestres";
            this.btnSemestres.UseVisualStyleBackColor = false;
            this.btnSemestres.Click += new System.EventHandler(this.btnSemestres_Click);
            // 
            // btnReglamentos
            // 
            this.btnReglamentos.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnReglamentos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReglamentos.Location = new System.Drawing.Point(145, 77);
            this.btnReglamentos.Name = "btnReglamentos";
            this.btnReglamentos.Size = new System.Drawing.Size(89, 39);
            this.btnReglamentos.TabIndex = 2;
            this.btnReglamentos.Text = "Reglamentos";
            this.btnReglamentos.UseVisualStyleBackColor = false;
            this.btnReglamentos.Click += new System.EventHandler(this.btnReglamentos_Click);
            // 
            // btnEstudiantesDoct
            // 
            this.btnEstudiantesDoct.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnEstudiantesDoct.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEstudiantesDoct.Location = new System.Drawing.Point(145, 122);
            this.btnEstudiantesDoct.Name = "btnEstudiantesDoct";
            this.btnEstudiantesDoct.Size = new System.Drawing.Size(89, 39);
            this.btnEstudiantesDoct.TabIndex = 2;
            this.btnEstudiantesDoct.Text = "Estudiantes Doctorado";
            this.btnEstudiantesDoct.UseVisualStyleBackColor = false;
            this.btnEstudiantesDoct.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnPropuestasDoct
            // 
            this.btnPropuestasDoct.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnPropuestasDoct.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPropuestasDoct.Location = new System.Drawing.Point(145, 167);
            this.btnPropuestasDoct.Name = "btnPropuestasDoct";
            this.btnPropuestasDoct.Size = new System.Drawing.Size(89, 39);
            this.btnPropuestasDoct.TabIndex = 2;
            this.btnPropuestasDoct.Text = "Propuestas Doctorado";
            this.btnPropuestasDoct.UseVisualStyleBackColor = false;
            this.btnPropuestasDoct.Click += new System.EventHandler(this.btnPropuestasDoct_Click);
            // 
            // HomeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(284, 282);
            this.ControlBox = false;
            this.Controls.Add(this.btnPropuestasDoct);
            this.Controls.Add(this.btnEstudiantesDoct);
            this.Controls.Add(this.btnReglamentos);
            this.Controls.Add(this.btnSemestres);
            this.Controls.Add(this.btnEscuelas);
            this.Controls.Add(this.btnProfesores);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "HomeForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BIENVENIDO A POSGRIQ";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnProfesores;
        private System.Windows.Forms.Button btnEscuelas;
        private System.Windows.Forms.Button btnSemestres;
        private System.Windows.Forms.Button btnReglamentos;
        private System.Windows.Forms.Button btnEstudiantesDoct;
        private System.Windows.Forms.Button btnPropuestasDoct;
    }
}