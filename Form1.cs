using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
namespace FinaleAppUtility
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PanelFunctions.ExcelPanelFunctions.OpenExcel();
            System.Windows.Forms.MessageBox.Show("Excel file was succesfuly updated");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(1);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Ing. Adrian Naziru is the pinguin master of this project");
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(PanelFunctions.ExcelPanelFunctions.ReadExcel(1, 2));
        }

        private void button7_Click(object sender, EventArgs e) //TEST
        {
            PanelFunctions.ExcelPanelFunctions.WriteExcelCell(1,2);
            System.Windows.Forms.MessageBox.Show("Excel file was succesfuly updated");
        }
    }
}
