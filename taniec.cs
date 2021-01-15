using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace WIZUALIZACJA_CAT_STREAM
{
    public partial class taniec : Form
    {
        private TextBox textBox1;
        private TextBox Spreadsheet;
        private ListBox dataGridView1;
        private Button button3;
        private OpenFileDialog openFileDialog1;

        public taniec()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)

        {
            listBox1.Font = new Font(FontFamily.GenericMonospace, textBox.Font.Size);
            listBox1.HorizontalScrollbar = true;
            ExcelHandler excelhandler;
            try
            {
                excelhandler = new ExcelHandler(openFileDialog1.FileName, Spreadsheet.Text);
            }
            catch
            {
                MessageBox.Show("Wybierz poprawny plik oraz wprowadź poprawną nazwę arkusza", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string text = textBox1.Text;
            if (String.IsNullOrEmpty(text))
            {
                MessageBox.Show("Wprowadź numery par", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            List<int> L = new List<int>();
            string temp = "";
            for (int i = 0; i < text.Length; i++)
            {
                if (Char.IsNumber(text[i])) temp += text[i];
                else if (!String.IsNullOrEmpty(temp))
                {
                    L.Add(int.Parse(temp) - 1);
                    temp = "";
                }
            }
            if (!String.IsNullOrEmpty(temp)) L.Add(int.Parse(temp) - 1);
            listBox1.Items.Clear();
            foreach (int element in L)
            {
                try
                {
                    listBox1.Items.Add(excelhandler.Row_toString(element));
                }
                catch
                {
                    MessageBox.Show("Para nie istnieje", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void InitializeComponent(taniec taniec)
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Spreadsheet = new System.Windows.Forms.TextBox();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(15, 59);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(165, 20);
            this.textBox1.TabIndex = 0;
            // 
            // Spreadsheet
            // 
            this.Spreadsheet.Location = new System.Drawing.Point(760, 91);
            this.Spreadsheet.Name = "Spreadsheet";
            this.Spreadsheet.Size = new System.Drawing.Size(194, 20);
            this.Spreadsheet.TabIndex = 1;
            this.Spreadsheet.Text = "Arkusz1";
            // 
            // listBox1
            // 
            this.listBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(136)))), ((int)(((byte)(228)))));
            this.listBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F);
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 31;
            this.listBox2.Items.AddRange(new object[] {
            "NUMERY PAR WRAZ Z NAZWISKAMI I KLUBEM"});
            this.listBox2.Location = new System.Drawing.Point(12, 91);
            this.listBox2.Name = "listBox1";
            this.listBox2.Size = new System.Drawing.Size(742, 407);
            this.listBox2.TabIndex = 2;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.Items.AddRange(new object[] {
            "Item 1, column 1",
            "Item 2, column 2",
            "Item 3, column 3",
            this.listBox1.Location = new System.Drawing.Point(0, 0);
            this.listBox1.MultiColumn = true;
            this.listBox1.Name = "listBox1";
            this.listBox1.ScrollAlwaysVisible = true;
            this.listBox1.Size = new System.Drawing.Size(120, 95);
            this.listBox1.TabIndex = 0;
            this.listBox1.ColumnWidth = 85;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(195, 59);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(139, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "pokaż";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(760, 16);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(194, 69);
            this.button2.TabIndex = 4;
            this.button2.Text = "WYBIERZ PLIK";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 12);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 5;
            this.button3.Text = "POWRÓT";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(966, 528);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.Spreadsheet);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "WIZUALIZACJA TANECZA BY BARTOSZ KOTEK";
            this.ResumeLayout(false);
            this.PerformLayout();
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Controls.Add(this.listBox1);
            this.Name = "Form1";
            this.ResumeLayout(false);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            HOME hOME = new HOME();
            hOME.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = new StreamWriter(@"lista.txt"))
            {

                foreach (String itemText in listBox2.Items)
                {

                    sw.Write(itemText);

                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelHandler excelhandler;
            try
            {
                excelhandler = new ExcelHandler(openFileDialog2.FileName, Spreadsheet2.Text);
            }
            catch
            {
                MessageBox.Show("Wybierz poprawny plik oraz wprowadź poprawną nazwę arkusza", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string text = textBox3.Text;
            if (String.IsNullOrEmpty(text))
            {
                MessageBox.Show("Wprowadź numer kategorii", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            List<int> L = new List<int>();
            string temp = "";
            for (int i = 0; i < text.Length; i++)
            {
                if (Char.IsNumber(text[i])) temp += text[i];
                else if (!String.IsNullOrEmpty(temp))
                {
                    L.Add(int.Parse(temp) - 1);
                    temp = "";
                }
            }
            if (!String.IsNullOrEmpty(temp)) L.Add(int.Parse(temp) - 1);
            listBox2.Items.Clear();
            foreach (int element in L)
            {
                try
                {
                    listBox2.Items.Add(excelhandler.Row_toString(element));
                }
                catch
                {
                    MessageBox.Show("Kategoria nie istnieje", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
        }
        
    }
}

        