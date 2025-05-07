using Exportar;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _6_de_Mayo
{
    public partial class Form1 : Form
       
    {
        Acciones acc = new Acciones();
        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnmostrar_Click(object sender, EventArgs e)
        {
            dgDatos.DataSource = acc.Mostrar();                                                                                                                                                                                                                                                                                                    
        }

        private void btnexportar_Click(object sender, EventArgs e)
        {
            if (acc.ExportaraExcel())
                MessageBox.Show("Exportado con exito...");
            else
                MessageBox.Show("Fallo el exportado...");
        }

        private void btnimportar_Click(object sender, EventArgs e)
        {
            
        }
    }
}
