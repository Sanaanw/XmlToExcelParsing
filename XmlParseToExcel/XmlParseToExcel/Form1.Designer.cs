using System;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Windows.Forms;
using XmlParseToExcel.Helpers;

namespace XmlParseToExcel
{
    partial class Form1 : Form
    {
        private System.ComponentModel.IContainer components = null;
        private Button button1;
        private Button button2;        
        private ListBox listBox1;      
        private ListBox listBox2;       
        private OpenFileDialog openFileDialog1;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "XML files (*.xml)|*.xml";
            openFileDialog1.Multiselect = true;

            if (openFileDialog1.ShowDialog() != DialogResult.OK)
                return;

            listBox1.Items.Clear();

            foreach (var filePath in openFileDialog1.FileNames)
            {
                listBox1.Items.Add(filePath);
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("Xaiş edirik XML faylı seçin.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                foreach (var item in listBox1.Items)
                {
                    string filePath = item.ToString();

                    var orderDto = XmlParserHelper.ParseXml(filePath);

                    string excelPath = @"C:\\Users\\ASUS\\OneDrive\\Desktop/Sipariş_Bravo-Market_2507024319615.xlsx";

                    ExcelExporterHelper.ExportXmlToExcel(orderDto, excelPath);

                    listBox2.Items.Add($"Exported: {Path.GetFileName(excelPath)}");
                }

                MessageBox.Show("Bütün fayllar uğurla Excel-ə çevrildi!", "Uğur", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xəta baş verdi: {ex.Message}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private Label label1;
        private Label label2;

        private void InitializeComponent()
        {
            button1 = new Button();
            button2 = new Button();
            listBox1 = new ListBox();
            listBox2 = new ListBox();
            label1 = new Label();
            label2 = new Label();
            openFileDialog1 = new OpenFileDialog();
            SuspendLayout();

            // button1
            button1.Location = new Point(400, 30);
            button1.Size = new Size(200, 50);
            button1.Name = "button1";
            button1.TabIndex = 0;
            button1.Text = "Faylları seçin.";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;

            // button2
            button2.Location = new Point(650, 30);
            button2.Size = new Size(200, 50);
            button2.Name = "button2";
            button2.TabIndex = 1;
            button2.Text = "İcra edin.";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;

            // label1
            label1.Location = new Point(50, 95);
            label1.Size = new Size(200, 20);
            label1.Text = "Seçilmiş fayllar:";
            label1.Font = new Font(label1.Font, FontStyle.Bold);

            // listBox1
            listBox1.FormattingEnabled = true;
            listBox1.Location = new Point(50, 120);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(500, 500);
            listBox1.TabIndex = 2;

            // label2
            label2.Location = new Point(600, 95);
            label2.Size = new Size(200, 20);
            label2.Text = "İcra edilmiş fayllar:";
            label2.Font = new Font(label2.Font, FontStyle.Bold);

            // listBox2 
            listBox2.FormattingEnabled = true;
            listBox2.Location = new Point(600, 120);
            listBox2.Name = "listBox2";
            listBox2.Size = new Size(500, 500);
            listBox2.TabIndex = 3;

            // openFileDialog1
            openFileDialog1.Filter = "XML files (*.xml)|*.xml";
            openFileDialog1.Multiselect = true;

            // Form1
            AutoScaleDimensions = new SizeF(8F, 16F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1200, 650);
            Controls.Add(button1);
            Controls.Add(button2);
            Controls.Add(label1);
            Controls.Add(listBox1);
            Controls.Add(label2);
            Controls.Add(listBox2);
            Name = "Form1";
            Text = "XML to Excel Parser";
            ResumeLayout(false);
        }


        #endregion
    }
}
