using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace onPoint
{
    public partial class UserControl1 : UserControl
    {
        Dictionary<int, SlideContents> SlideData;
        volatile int currentKey = 0;
        volatile String key = "";
        public UserControl1(Dictionary<int, SlideContents> d)
        {
            InitializeComponent();
            SlideData = d;
        }
        public void changeSlide(int id)
        {
            currentKey = id;
            switchData();
        }

        private void switchData()
        { 
             textBox1.Text = SlideData[currentKey].name;
             numericUpDown1.Value = SlideData[currentKey].time; 
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SlideData[currentKey].name = textBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            SlideData[currentKey].time = (int)numericUpDown1.Value;
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            key = textBox2.Text;
        }
    }
}
