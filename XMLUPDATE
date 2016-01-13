using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace UpdateRTUEMuXMLFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.Visible = false;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDown;
        }

        private string _fname;
        public string fname
        {
            get { return _fname; }
            set { _fname = value; }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // Browse button
            openFileDialog1.ShowDialog();
            string flname = openFileDialog1.FileName;
            label1.Text = "File Opened: "+flname;
            this.fname = flname;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string newValue = string.Empty;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(this.fname);
            XmlNodeList xnList = null;
            xnList = xmlDoc.GetElementsByTagName(comboBox1.SelectedItem.ToString());
            foreach (XmlNode xn in xnList)
            {
                string attributesvalues = xn.Attributes[comboBox2.SelectedItem.ToString()].InnerText;
                if (comboBox2.SelectedItem.ToString() == "Other(Please Specify)")
                {
                    textBox1.Visible = true;
                    attributesvalues = xn.Attributes[textBox1.Text].InnerText;

                }
                if (attributesvalues == textBox3.Text)
                {
                    xn.Attributes[comboBox3.SelectedItem.ToString()].Value = textBox2.Text;
                }
            }
            xmlDoc.Save(this.fname);
            label1.Text  = "Updated the Regsiter "+textBox1.Text+"with value = "+textBox2.Text;
        }
    }
}
