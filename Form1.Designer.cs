namespace _6_de_Mayo
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgDatos = new System.Windows.Forms.DataGridView();
            this.btnmostrar = new System.Windows.Forms.Button();
            this.btnexportar = new System.Windows.Forms.Button();
            this.btnimportar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgDatos)).BeginInit();
            this.SuspendLayout();
            // 
            // dgDatos
            // 
            this.dgDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgDatos.Location = new System.Drawing.Point(70, 58);
            this.dgDatos.Name = "dgDatos";
            this.dgDatos.Size = new System.Drawing.Size(658, 170);
            this.dgDatos.TabIndex = 0;
            this.dgDatos.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // btnmostrar
            // 
            this.btnmostrar.Location = new System.Drawing.Point(70, 278);
            this.btnmostrar.Name = "btnmostrar";
            this.btnmostrar.Size = new System.Drawing.Size(94, 30);
            this.btnmostrar.TabIndex = 1;
            this.btnmostrar.Text = "Mostrar";
            this.btnmostrar.UseVisualStyleBackColor = true;
            this.btnmostrar.Click += new System.EventHandler(this.btnmostrar_Click);
            // 
            // btnexportar
            // 
            this.btnexportar.Location = new System.Drawing.Point(634, 278);
            this.btnexportar.Name = "btnexportar";
            this.btnexportar.Size = new System.Drawing.Size(94, 30);
            this.btnexportar.TabIndex = 2;
            this.btnexportar.Text = "Exportar";
            this.btnexportar.UseVisualStyleBackColor = true;
            this.btnexportar.Click += new System.EventHandler(this.btnexportar_Click);
            // 
            // btnimportar
            // 
            this.btnimportar.Location = new System.Drawing.Point(358, 282);
            this.btnimportar.Name = "btnimportar";
            this.btnimportar.Size = new System.Drawing.Size(75, 23);
            this.btnimportar.TabIndex = 3;
            this.btnimportar.Text = "Importar";
            this.btnimportar.UseVisualStyleBackColor = true;
            this.btnimportar.Click += new System.EventHandler(this.btnimportar_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnimportar);
            this.Controls.Add(this.btnexportar);
            this.Controls.Add(this.btnmostrar);
            this.Controls.Add(this.dgDatos);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgDatos)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgDatos;
        private System.Windows.Forms.Button btnmostrar;
        private System.Windows.Forms.Button btnexportar;
        private System.Windows.Forms.Button btnimportar;
    }
}

