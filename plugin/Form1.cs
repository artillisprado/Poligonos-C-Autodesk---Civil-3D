using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace plugin
{
    public partial class Form1 : Form
    {
        Class1 class1 = new Class1();
        string filePath;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //class1.readExcel("C:\\Users\\artillis.prado\\Downloads\\NSPT solido.xlsx");
            using(OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.Filter = "Planilha do Microsoft Excel (.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog1.Title = "Localizar Arquivos";
                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog1.FileName;
                    textBox1.Text = filePath;
                    //class1.readExcel(filePath);
                }
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        { }

        private void button2_Click(object sender, EventArgs e)
        {
            int diametro = Convert.ToInt32(Math.Round(numericUpDown1.Value, 2));
            this.Close();
            class1.readExcel(filePath, diametro);

        }

        private void label1_Click(object sender, EventArgs e)
        { }

        private void label2_Click(object sender, EventArgs e)
        { }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        { }
    }
}
