namespace PosgrIQ
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.menuMain = new System.Windows.Forms.MenuStrip();
            this.menuArchivo = new System.Windows.Forms.ToolStripMenuItem();
            this.menuConfiguracion = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSalir = new System.Windows.Forms.ToolStripMenuItem();
            this.tablasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuDoct = new System.Windows.Forms.ToolStripMenuItem();
            this.menuEstudiantesDoct = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMatriculaDoct = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPonenciasDoct = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPropuestasDoct = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPublicacionesDoct = new System.Windows.Forms.ToolStripMenuItem();
            this.menuTesisDoct = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMaes = new System.Windows.Forms.ToolStripMenuItem();
            this.menuEstudiantesMaes = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMatriculaMaes = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPonenciasMaes = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPropuestasMaes = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPublicacionesMaes = new System.Windows.Forms.ToolStripMenuItem();
            this.menuTesisMaes = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.escuelasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.profesoresToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reglamentosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.semestresToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ventanasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cerrarTodo = new System.Windows.Forms.ToolStripMenuItem();
            this.menuAcerca = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuMain
            // 
            this.menuMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuArchivo,
            this.tablasToolStripMenuItem,
            this.ventanasToolStripMenuItem,
            this.menuAcerca});
            this.menuMain.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.menuMain.Location = new System.Drawing.Point(0, 0);
            this.menuMain.MdiWindowListItem = this.ventanasToolStripMenuItem;
            this.menuMain.Name = "menuMain";
            this.menuMain.Size = new System.Drawing.Size(1008, 24);
            this.menuMain.TabIndex = 1;
            this.menuMain.Text = "menuStrip1";
            // 
            // menuArchivo
            // 
            this.menuArchivo.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuConfiguracion,
            this.menuSalir});
            this.menuArchivo.Name = "menuArchivo";
            this.menuArchivo.Size = new System.Drawing.Size(60, 20);
            this.menuArchivo.Text = "&Archivo";
            // 
            // menuConfiguracion
            // 
            this.menuConfiguracion.Name = "menuConfiguracion";
            this.menuConfiguracion.Size = new System.Drawing.Size(150, 22);
            this.menuConfiguracion.Text = "&Configuración";
            this.menuConfiguracion.Click += new System.EventHandler(this.menuConfiguracion_Click);
            // 
            // menuSalir
            // 
            this.menuSalir.Name = "menuSalir";
            this.menuSalir.Size = new System.Drawing.Size(150, 22);
            this.menuSalir.Text = "&Salir";
            this.menuSalir.Click += new System.EventHandler(this.salirToolStripMenuItem_Click);
            // 
            // tablasToolStripMenuItem
            // 
            this.tablasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuDoct,
            this.menuMaes,
            this.toolStripSeparator1,
            this.escuelasToolStripMenuItem,
            this.profesoresToolStripMenuItem,
            this.reglamentosToolStripMenuItem,
            this.semestresToolStripMenuItem});
            this.tablasToolStripMenuItem.Name = "tablasToolStripMenuItem";
            this.tablasToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.tablasToolStripMenuItem.Text = "Tablas";
            // 
            // menuDoct
            // 
            this.menuDoct.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuEstudiantesDoct,
            this.menuMatriculaDoct,
            this.menuPonenciasDoct,
            this.menuPropuestasDoct,
            this.menuPublicacionesDoct,
            this.menuTesisDoct});
            this.menuDoct.Name = "menuDoct";
            this.menuDoct.Size = new System.Drawing.Size(143, 22);
            this.menuDoct.Text = "&Doctorado";
            // 
            // menuEstudiantesDoct
            // 
            this.menuEstudiantesDoct.Name = "menuEstudiantesDoct";
            this.menuEstudiantesDoct.Size = new System.Drawing.Size(147, 22);
            this.menuEstudiantesDoct.Text = "&Estudiantes";
            this.menuEstudiantesDoct.Click += new System.EventHandler(this.menuEstudiantesDoct_Click);
            // 
            // menuMatriculaDoct
            // 
            this.menuMatriculaDoct.Name = "menuMatriculaDoct";
            this.menuMatriculaDoct.Size = new System.Drawing.Size(147, 22);
            this.menuMatriculaDoct.Text = "&Matrícula";
            this.menuMatriculaDoct.Click += new System.EventHandler(this.menuMatriculaDoct_Click);
            // 
            // menuPonenciasDoct
            // 
            this.menuPonenciasDoct.Name = "menuPonenciasDoct";
            this.menuPonenciasDoct.Size = new System.Drawing.Size(147, 22);
            this.menuPonenciasDoct.Text = "&Ponencias";
            this.menuPonenciasDoct.Click += new System.EventHandler(this.menuPonenciasDoct_Click);
            // 
            // menuPropuestasDoct
            // 
            this.menuPropuestasDoct.Name = "menuPropuestasDoct";
            this.menuPropuestasDoct.Size = new System.Drawing.Size(147, 22);
            this.menuPropuestasDoct.Text = "P&ropuestas";
            this.menuPropuestasDoct.Click += new System.EventHandler(this.menuPropuestasDoct_Click);
            // 
            // menuPublicacionesDoct
            // 
            this.menuPublicacionesDoct.Name = "menuPublicacionesDoct";
            this.menuPublicacionesDoct.Size = new System.Drawing.Size(147, 22);
            this.menuPublicacionesDoct.Text = "P&ublicaciones";
            this.menuPublicacionesDoct.Click += new System.EventHandler(this.menuPublicacionesDoct_Click);
            // 
            // menuTesisDoct
            // 
            this.menuTesisDoct.Name = "menuTesisDoct";
            this.menuTesisDoct.Size = new System.Drawing.Size(147, 22);
            this.menuTesisDoct.Text = "&Tesis";
            this.menuTesisDoct.Click += new System.EventHandler(this.menuTesisDoct_Click);
            // 
            // menuMaes
            // 
            this.menuMaes.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuEstudiantesMaes,
            this.menuMatriculaMaes,
            this.menuPonenciasMaes,
            this.menuPropuestasMaes,
            this.menuPublicacionesMaes,
            this.menuTesisMaes});
            this.menuMaes.Name = "menuMaes";
            this.menuMaes.Size = new System.Drawing.Size(143, 22);
            this.menuMaes.Text = "&Maestría";
            // 
            // menuEstudiantesMaes
            // 
            this.menuEstudiantesMaes.Name = "menuEstudiantesMaes";
            this.menuEstudiantesMaes.Size = new System.Drawing.Size(147, 22);
            this.menuEstudiantesMaes.Text = "&Estudiantes";
            this.menuEstudiantesMaes.Click += new System.EventHandler(this.menuEstudiantesMaes_Click);
            // 
            // menuMatriculaMaes
            // 
            this.menuMatriculaMaes.Name = "menuMatriculaMaes";
            this.menuMatriculaMaes.Size = new System.Drawing.Size(147, 22);
            this.menuMatriculaMaes.Text = "&Matrícula";
            this.menuMatriculaMaes.Click += new System.EventHandler(this.menuMatriculaMaes_Click);
            // 
            // menuPonenciasMaes
            // 
            this.menuPonenciasMaes.Name = "menuPonenciasMaes";
            this.menuPonenciasMaes.Size = new System.Drawing.Size(147, 22);
            this.menuPonenciasMaes.Text = "&Ponencias";
            this.menuPonenciasMaes.Click += new System.EventHandler(this.menuPonenciasMaes_Click);
            // 
            // menuPropuestasMaes
            // 
            this.menuPropuestasMaes.Name = "menuPropuestasMaes";
            this.menuPropuestasMaes.Size = new System.Drawing.Size(147, 22);
            this.menuPropuestasMaes.Text = "P&ropuestas";
            this.menuPropuestasMaes.Click += new System.EventHandler(this.menuPropuestasMaes_Click);
            // 
            // menuPublicacionesMaes
            // 
            this.menuPublicacionesMaes.Name = "menuPublicacionesMaes";
            this.menuPublicacionesMaes.Size = new System.Drawing.Size(147, 22);
            this.menuPublicacionesMaes.Text = "P&ublicaciones";
            this.menuPublicacionesMaes.Click += new System.EventHandler(this.menuPublicacionesMaes_Click);
            // 
            // menuTesisMaes
            // 
            this.menuTesisMaes.Name = "menuTesisMaes";
            this.menuTesisMaes.Size = new System.Drawing.Size(147, 22);
            this.menuTesisMaes.Text = "&Tesis";
            this.menuTesisMaes.Click += new System.EventHandler(this.menuTesisMaes_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(140, 6);
            // 
            // escuelasToolStripMenuItem
            // 
            this.escuelasToolStripMenuItem.Name = "escuelasToolStripMenuItem";
            this.escuelasToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.escuelasToolStripMenuItem.Text = "&Escuelas";
            this.escuelasToolStripMenuItem.Click += new System.EventHandler(this.escuelasToolStripMenuItem_Click);
            // 
            // profesoresToolStripMenuItem
            // 
            this.profesoresToolStripMenuItem.Name = "profesoresToolStripMenuItem";
            this.profesoresToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.profesoresToolStripMenuItem.Text = "&Profesores";
            this.profesoresToolStripMenuItem.Click += new System.EventHandler(this.profesoresToolStripMenuItem_Click);
            // 
            // reglamentosToolStripMenuItem
            // 
            this.reglamentosToolStripMenuItem.Name = "reglamentosToolStripMenuItem";
            this.reglamentosToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.reglamentosToolStripMenuItem.Text = "&Reglamentos";
            this.reglamentosToolStripMenuItem.Click += new System.EventHandler(this.reglamentosToolStripMenuItem_Click);
            // 
            // semestresToolStripMenuItem
            // 
            this.semestresToolStripMenuItem.Name = "semestresToolStripMenuItem";
            this.semestresToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
            this.semestresToolStripMenuItem.Text = "&Semestres";
            this.semestresToolStripMenuItem.Click += new System.EventHandler(this.semestresToolStripMenuItem_Click);
            // 
            // ventanasToolStripMenuItem
            // 
            this.ventanasToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cerrarTodo});
            this.ventanasToolStripMenuItem.Name = "ventanasToolStripMenuItem";
            this.ventanasToolStripMenuItem.Size = new System.Drawing.Size(67, 20);
            this.ventanasToolStripMenuItem.Text = "&Ventanas";
            // 
            // cerrarTodo
            // 
            this.cerrarTodo.Name = "cerrarTodo";
            this.cerrarTodo.Size = new System.Drawing.Size(134, 22);
            this.cerrarTodo.Text = "&Cerrar todo";
            this.cerrarTodo.Click += new System.EventHandler(this.cerrarTodoToolStripMenuItem_Click);
            // 
            // menuAcerca
            // 
            this.menuAcerca.Name = "menuAcerca";
            this.menuAcerca.Size = new System.Drawing.Size(71, 20);
            this.menuAcerca.Text = "&Acerca de";
            this.menuAcerca.Click += new System.EventHandler(this.menuAcerca_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(1008, 731);
            this.Controls.Add(this.menuMain);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuMain;
            this.Name = "MainForm";
            this.Text = "POSGRIQ: ADMINISTRADOR DE INFORMACION DE POSGRADO DE INQ QUIMICA";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuMain.ResumeLayout(false);
            this.menuMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuMain;
        private System.Windows.Forms.ToolStripMenuItem menuArchivo;
        private System.Windows.Forms.ToolStripMenuItem menuConfiguracion;
        private System.Windows.Forms.ToolStripMenuItem menuSalir;
        private System.Windows.Forms.ToolStripMenuItem tablasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem menuDoct;
        private System.Windows.Forms.ToolStripMenuItem menuEstudiantesDoct;
        private System.Windows.Forms.ToolStripMenuItem menuMatriculaDoct;
        private System.Windows.Forms.ToolStripMenuItem menuPonenciasDoct;
        private System.Windows.Forms.ToolStripMenuItem menuPropuestasDoct;
        private System.Windows.Forms.ToolStripMenuItem menuPublicacionesDoct;
        private System.Windows.Forms.ToolStripMenuItem menuTesisDoct;
        private System.Windows.Forms.ToolStripMenuItem menuMaes;
        private System.Windows.Forms.ToolStripMenuItem menuEstudiantesMaes;
        private System.Windows.Forms.ToolStripMenuItem menuMatriculaMaes;
        private System.Windows.Forms.ToolStripMenuItem menuPonenciasMaes;
        private System.Windows.Forms.ToolStripMenuItem menuPropuestasMaes;
        private System.Windows.Forms.ToolStripMenuItem menuPublicacionesMaes;
        private System.Windows.Forms.ToolStripMenuItem menuTesisMaes;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem escuelasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem profesoresToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reglamentosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem semestresToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ventanasToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cerrarTodo;
        private System.Windows.Forms.ToolStripMenuItem menuAcerca;

    }
}

