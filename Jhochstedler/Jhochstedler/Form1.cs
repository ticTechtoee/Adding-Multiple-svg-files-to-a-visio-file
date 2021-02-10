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
using Microsoft.Office.Interop.Visio;


namespace Jhochstedler
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            try

            {
                Microsoft.Office.Interop.Visio.Application app = new Microsoft.Office.Interop.Visio.Application();




                Document doc = app.Documents.Add(pathVisio);

                Page page;
                Page page1 = doc.Pages[1];


                string path = pathSvg;
                string[] files = Directory.GetFiles(path);


                string[] fileName = Directory.GetFiles(path).Select(file => System.IO.Path.GetFileName(file)).ToArray();


                int f = 1;
                int n;

                page1.Name = fileName[0].ToString();
                doc.Pages[1].Import(files[0]);
                doc.Pages[1].ResizeToFitContents();

                for (n = 2; n <= files.Length; n++)
                {


                    page = doc.Pages.Add();

                    page.Name = fileName[f].ToString();

                    doc.Pages[n].Import(files[f]);
                    doc.Pages[n].ResizeToFitContents();
                    
                    f++;
                    
                }

                doc.PrintFitOnPages = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("We have got this error. Please Contact Developer" + ex,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }











        }
        string pathVisio;
        string pathSvg;

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog slctfile = new OpenFileDialog();
            slctfile.Title = "Please Select Your Visio File";
            if (slctfile.ShowDialog() == DialogResult.OK)
            {
                pathVisio = slctfile.FileName;
                label3.Text = pathVisio;
                button2.TextAlign = ContentAlignment.TopCenter;
                button2.Image = Properties.Resources.checked_symbol__1_;
                
                button2.ImageAlign = ContentAlignment.MiddleRight;
             
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog slctfile = new FolderBrowserDialog();
            slctfile.Description = "Please Select Your SVG File Folder";
            if (slctfile.ShowDialog() == DialogResult.OK)
            {
                pathSvg = slctfile.SelectedPath;

                label4.Text = pathSvg;
                button1.TextAlign = ContentAlignment.TopCenter;
                button1.Image = Properties.Resources.checked_symbol__1_;
                button1.ImageAlign = ContentAlignment.MiddleRight;

            }
        }
    }
}
