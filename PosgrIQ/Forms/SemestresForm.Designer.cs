namespace PosgrIQ
{
    partial class SemestresForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SemestresForm));
            this.dataGridSemestres = new System.Windows.Forms.DataGridView();
            this.btnMod = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnCerrar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSemestres)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridSemestres
            // 
            this.dataGridSemestres.AllowUserToAddRows = false;
            this.dataGridSemestres.AllowUserToDeleteRows = false;
            this.dataGridSemestres.AllowUserToResizeRows = false;
            this.dataGridSemestres.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridSemestres.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridSemestres.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridSemestres.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridSemestres.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridSemestres.GridColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridSemestres.Location = new System.Drawing.Point(10, 10);
            this.dataGridSemestres.MultiSelect = false;
            this.dataGridSemestres.Name = "dataGridSemestres";
            this.dataGridSemestres.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridSemestres.Size = new System.Drawing.Size(800, 200);
            this.dataGridSemestres.TabIndex = 1;
            // 
            // btnMod
            // 
            this.btnMod.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnMod.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMod.Location = new System.Drawing.Point(653, 216);
            this.btnMod.Name = "btnMod";
            this.btnMod.Size = new System.Drawing.Size(75, 23);
            this.btnMod.TabIndex = 5;
            this.btnMod.Text = "Modificar";
            this.btnMod.UseVisualStyleBackColor = false;
            this.btnMod.Click += new System.EventHandler(this.btnMod_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdd.Location = new System.Drawing.Point(571, 216);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 4;
            this.btnAdd.Text = "Agregar";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnCerrar
            // 
            this.btnCerrar.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCerrar.Location = new System.Drawing.Point(735, 216);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(75, 23);
            this.btnCerrar.TabIndex = 3;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // SemestresForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(817, 247);
            this.ControlBox = false;
            this.Controls.Add(this.btnMod);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.dataGridSemestres);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SemestresForm";
            this.Text = "SEMESTRES";
            this.Load += new System.EventHandler(this.SemestresForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSemestres)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridSemestres;
        private System.Windows.Forms.Button btnMod;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnCerrar;
    }
}