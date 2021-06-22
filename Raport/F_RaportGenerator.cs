using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Raport
{
    public partial class F_RaportGenerator : Form
    {
        private string filePath, fileContent;

        public F_RaportGenerator()
        {
            InitializeComponent();
        }

        private void portycsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                filePath = openFileDialog1.FileName;

                //Read the contents of the file into a stream
                var fileStream = openFileDialog1.OpenFile();
                string linijka;

                using (StreamReader reader = new StreamReader(fileStream))
                {
                    //fileContent = reader.ReadToEnd();
                    while((linijka = reader.ReadLine()) != null)
                    {
                        linijka = linijka.Trim(new Char[] { ' ', '"', '.' });
                        fileContent += linijka.ToString();
                    }                    
                }

                //usuwam spacje
                //fileContent = fileContent.Trim(' ', '"');

                MessageBox.Show(fileContent);
            }
        }

        private void WalidujMetasploitPortyCSV() 
        {
            //services -u -c name,proto -o /home/user/export_wynik.csv

        }

    }
}
