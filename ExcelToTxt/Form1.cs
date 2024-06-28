using System;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace ExcelColumnViewer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int sheetNumber = (int)numericUpDown1.Value;
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Please enter the path to the Excel file.");
                return;
            }

            try
            {
                // Load the Excel file
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(textBox1.Text);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[sheetNumber];

                // Clear the ListBox before adding new items
                listBox1.Items.Clear();

                // Get the column names from the first row
                for (int j = 1; j <= sheet.Columns.Length; j++)
                {
                    string columnName = sheet.Range[1, j].Text;
                    listBox1.Items.Add(columnName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Excel file: " + ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int sheetNumber = (int)numericUpDown1.Value;
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Please enter the path to the Excel file and the format text.");
                return;
            }

            try
            {
                // Načítanie Excel súboru
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(textBox1.Text);

                Worksheet sheet = workbook.Worksheets[sheetNumber];

                // Vytvorenie StringBuildera pre generovaný text
                StringBuilder sb = new StringBuilder();

                // Získanie šablóny z textBox2
                string template = textBox2.Text;

                // Prechod cez každý riadok v hárku
                for (int i = 2; i <= sheet.Rows.Length; i++) // Dáta začínajú od druhého riadku
                {
                    string rowText = template;

                    // Prechod cez každý stĺpec v riadku
                    for (int j = 1; j <= sheet.Columns.Length; j++)
                    {
                        string columnName = sheet.Range[1, j].Text;
                        string cellValue = sheet.Range[i, j].Value;

                        // Ak ide o hodnotu v prvom stĺpci, upravíme ju na 5-miestne číslo
                        if (j == 1)
                        {
                            cellValue = FormatToFiveDigits(cellValue);
                        }

                        rowText = rowText.Replace($"@{columnName}@", cellValue);
                    }

                    // Pridanie generovaného riadku textu do StringBuildera
                    sb.AppendLine(rowText);
                }

                // Uloženie generovaného textu do súboru
                File.WriteAllText("GeneratedText.txt", sb.ToString(), Encoding.UTF8);

                // Notifikácia používateľa o dokončení procesu
                MessageBox.Show("Generation completed successfully.");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error generating text file: " + ex.Message);
            }
        }

        private string FormatToFiveDigits(string value)
        {
            // Odstrániť prázdne znaky a pokúsiť sa konvertovať na číslo
            value = value.Trim();
            if (int.TryParse(value, out int numericValue))
            {
                return numericValue.ToString("D5"); // Konvertujeme na 5-miestne číslo
            }
            else
            {
                // Ak sa nedá konvertovať, vrátime pôvodnú hodnotu alebo zhodíme výnimku, podľa potreby
                return value;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}
