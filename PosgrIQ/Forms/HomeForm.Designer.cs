﻿namespace PosgrIQ
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
            this.SuspendLayout();
            // 
            // btnProfesores
            // 
            this.btnProfesores.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnProfesores.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnProfesores.Location = new System.Drawing.Point(108, 119);
            this.btnProfesores.Name = "btnProfesores";
            this.btnProfesores.Size = new System.Drawing.Size(75, 23);
            this.btnProfesores.TabIndex = 0;
            this.btnProfesores.Text = "Profesores";
            this.btnProfesores.UseVisualStyleBackColor = false;
            this.btnProfesores.Click += new System.EventHandler(this.btnProfesores_Click);
            // 
            // btnEscuelas
            // 
            this.btnEscuelas.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnEscuelas.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEscuelas.Location = new System.Drawing.Point(108, 148);
            this.btnEscuelas.Name = "btnEscuelas";
            this.btnEscuelas.Size = new System.Drawing.Size(75, 23);
            this.btnEscuelas.TabIndex = 1;
            this.btnEscuelas.Text = "Escuelas";
            this.btnEscuelas.UseVisualStyleBackColor = false;
            this.btnEscuelas.Click += new System.EventHandler(this.btnEscuelas_Click);
            // 
            // HomeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(284, 282);
            this.ControlBox = false;
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
    }
}