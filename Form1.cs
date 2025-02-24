using System;
using System.Windows.Forms;

namespace Specification
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) { }

        private async void button1_Click(object sender, EventArgs e)
        {
            string sectionNumber = CommonUtilities.PromptSectionNumber(textBox1);
            if (string.IsNullOrEmpty(sectionNumber)) return;
            string projectNumber = textBox2.Text;

            button1.Enabled = false;
            progressBar1.Value = 0;
            var progress = new Progress<int>(percent => progressBar1.Value = percent);

            await Task.Run(() => SpecificationProcessor.ProcessSpecification(sectionNumber, projectNumber, progress));

            button1.Enabled = true;
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            string sectionNumber = CommonUtilities.PromptSectionNumber(textBox1);
            if (string.IsNullOrEmpty(sectionNumber)) return;
            string projectNumber = textBox2.Text;

            button2.Enabled = false;
            progressBar1.Value = 0;
            var progress = new Progress<int>(percent => progressBar1.Value = percent);

            await Task.Run(() => MaterialDataProcessor.ProcessMaterialData(sectionNumber, projectNumber, progress));

            button2.Enabled = true;
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            string sectionNumber = CommonUtilities.PromptSectionNumber(textBox1);
            if (string.IsNullOrEmpty(sectionNumber)) return;
            string projectNumber = textBox2.Text;

            button3.Enabled = false;
            progressBar1.Value = 0;
            var progress = new Progress<int>(percent => progressBar1.Value = percent);

            await Task.Run(() => UnitsProcessor.ProcessUnits(sectionNumber, projectNumber, progress));

            button3.Enabled = true;
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;
            progressBar1.Value = 0;
            var progress = new Progress<int>(percent => progressBar1.Value = percent);

            await Task.Run(() => ClearExtraSheetsProcessor.ClearSheets(progress));

            button4.Enabled = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void textBox2_TextChanged(object sender, EventArgs e) { }
        private void label1_Click(object sender, EventArgs e) { }
        private void textBox1_TextChanged_1(object sender, EventArgs e) { }
    }
}