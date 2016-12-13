namespace PosgrIQ
{
    partial class ConfiguracionForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfiguracionForm));
            this.label1 = new System.Windows.Forms.Label();
            this.txtCorreo = new System.Windows.Forms.TextBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtClave = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtRutaOne = new System.Windows.Forms.TextBox();
            this.btnRutaOne = new System.Windows.Forms.Button();
            this.chkCambios = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtRutaBD = new System.Windows.Forms.TextBox();
            this.btnRutaBD = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txtDirector = new System.Windows.Forms.TextBox();
            this.txtCoordinador = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 14);
            this.label1.TabIndex = 0;
            this.label1.Text = "Correo:";
            // 
            // txtCorreo
            // 
            this.txtCorreo.Enabled = false;
            this.txtCorreo.Location = new System.Drawing.Point(98, 42);
            this.txtCorreo.Name = "txtCorreo";
            this.txtCorreo.Size = new System.Drawing.Size(212, 22);
            this.txtCorreo.TabIndex = 2;
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnClose.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClose.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnClose.Location = new System.Drawing.Point(553, 241);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 8;
            this.btnClose.Text = "Cerrar";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 14);
            this.label2.TabIndex = 0;
            this.label2.Text = "Clave:";
            // 
            // txtClave
            // 
            this.txtClave.Enabled = false;
            this.txtClave.Location = new System.Drawing.Point(98, 70);
            this.txtClave.Name = "txtClave";
            this.txtClave.PasswordChar = '*';
            this.txtClave.Size = new System.Drawing.Size(212, 22);
            this.txtClave.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 160);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 14);
            this.label3.TabIndex = 0;
            this.label3.Text = "Ruta OneDrive:";
            // 
            // txtRutaOne
            // 
            this.txtRutaOne.Enabled = false;
            this.txtRutaOne.Location = new System.Drawing.Point(98, 155);
            this.txtRutaOne.Name = "txtRutaOne";
            this.txtRutaOne.Size = new System.Drawing.Size(530, 22);
            this.txtRutaOne.TabIndex = 5;
            this.txtRutaOne.Text = "C:\\Users\\Hdsp\\OneDrive\\OneDrive - Universidad Industrial de Santander";
            // 
            // btnRutaOne
            // 
            this.btnRutaOne.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnRutaOne.Enabled = false;
            this.btnRutaOne.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnRutaOne.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRutaOne.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnRutaOne.Location = new System.Drawing.Point(98, 183);
            this.btnRutaOne.Name = "btnRutaOne";
            this.btnRutaOne.Size = new System.Drawing.Size(107, 23);
            this.btnRutaOne.TabIndex = 25;
            this.btnRutaOne.Text = "Seleccionar ruta";
            this.btnRutaOne.UseVisualStyleBackColor = false;
            this.btnRutaOne.Click += new System.EventHandler(this.btnRutaOne_Click);
            // 
            // chkCambios
            // 
            this.chkCambios.AutoSize = true;
            this.chkCambios.Location = new System.Drawing.Point(7, 12);
            this.chkCambios.Name = "chkCambios";
            this.chkCambios.Size = new System.Drawing.Size(124, 18);
            this.chkCambios.TabIndex = 1;
            this.chkCambios.Text = "Habilitar cambios";
            this.chkCambios.UseVisualStyleBackColor = true;
            this.chkCambios.CheckedChanged += new System.EventHandler(this.chkCambios_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 102);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 14);
            this.label4.TabIndex = 0;
            this.label4.Text = "Ruta BD:";
            // 
            // txtRutaBD
            // 
            this.txtRutaBD.Enabled = false;
            this.txtRutaBD.Location = new System.Drawing.Point(98, 98);
            this.txtRutaBD.Name = "txtRutaBD";
            this.txtRutaBD.Size = new System.Drawing.Size(530, 22);
            this.txtRutaBD.TabIndex = 4;
            this.txtRutaBD.Text = "C:\\Users\\Hdsp\\OneDrive\\OneDrive - Universidad Industrial de Santander";
            // 
            // btnRutaBD
            // 
            this.btnRutaBD.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnRutaBD.Enabled = false;
            this.btnRutaBD.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnRutaBD.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRutaBD.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnRutaBD.Location = new System.Drawing.Point(98, 126);
            this.btnRutaBD.Name = "btnRutaBD";
            this.btnRutaBD.Size = new System.Drawing.Size(107, 23);
            this.btnRutaBD.TabIndex = 25;
            this.btnRutaBD.Text = "Seleccionar ruta";
            this.btnRutaBD.UseVisualStyleBackColor = false;
            this.btnRutaBD.Click += new System.EventHandler(this.btnRutaBD_Click);
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(4, 211);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(88, 32);
            this.label5.TabIndex = 0;
            this.label5.Text = "Director de Escuela:";
            // 
            // txtDirector
            // 
            this.txtDirector.Enabled = false;
            this.txtDirector.Location = new System.Drawing.Point(98, 212);
            this.txtDirector.Name = "txtDirector";
            this.txtDirector.Size = new System.Drawing.Size(200, 22);
            this.txtDirector.TabIndex = 6;
            this.txtDirector.Text = "C:\\Users\\Hdsp\\OneDrive\\OneDrive - Universidad Industrial de Santander";
            // 
            // txtCoordinador
            // 
            this.txtCoordinador.Enabled = false;
            this.txtCoordinador.Location = new System.Drawing.Point(428, 211);
            this.txtCoordinador.Name = "txtCoordinador";
            this.txtCoordinador.Size = new System.Drawing.Size(200, 22);
            this.txtCoordinador.TabIndex = 7;
            this.txtCoordinador.Text = "C:\\Users\\Hdsp\\OneDrive\\OneDrive - Universidad Industrial de Santander";
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(339, 211);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(88, 32);
            this.label6.TabIndex = 27;
            this.label6.Text = "Coordinador de Posgrado:";
            // 
            // ConfiguracionForm
            // 
            this.AcceptButton = this.btnClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(635, 274);
            this.ControlBox = false;
            this.Controls.Add(this.txtCoordinador);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.chkCambios);
            this.Controls.Add(this.btnRutaBD);
            this.Controls.Add(this.btnRutaOne);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.txtRutaBD);
            this.Controls.Add(this.txtDirector);
            this.Controls.Add(this.txtRutaOne);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtClave);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCorreo);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ConfiguracionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CONFIGURACION POSGRIQ";
            this.Load += new System.EventHandler(this.ConfiguracionForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCorreo;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtClave;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtRutaOne;
        private System.Windows.Forms.Button btnRutaOne;
        private System.Windows.Forms.CheckBox chkCambios;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtRutaBD;
        private System.Windows.Forms.Button btnRutaBD;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtDirector;
        private System.Windows.Forms.TextBox txtCoordinador;
        private System.Windows.Forms.Label label6;
    }
}