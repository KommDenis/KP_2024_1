using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;

namespace CopyrightRegistration
{
    public partial class Form1 : Form
    {
        private CopyrightRegistrationManager manager;
        private ToolTip toolTip;

        public Form1()
        {
            InitializeComponent();
            InitializeForm();
        }

        private void InitializeForm()
        {
            manager = new CopyrightRegistrationManager();
            toolTip = new ToolTip();

            RefreshText();
            ApplyCustomStyles();
            ApplyToolTips();

            btnAdd.Click += BtnAdd_Click;
            btnExport.Click += BtnExport_Click;
   
            btnRefresh.Click += BtnRefresh_Click;
        }

        private void ApplyCustomStyles()
        {
            this.BackColor = Color.LightGray;
            this.Font = new Font("Arial", 10F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

            ApplyButtonStyle(btnAdd);
            ApplyButtonStyle(btnExport);
         
            ApplyButtonStyle(btnRefresh);
        }

        private void ApplyButtonStyle(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = Color.DarkCyan;
            button.ForeColor = Color.White;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.Teal;
            button.FlatAppearance.MouseDownBackColor = Color.DarkSlateGray;
            button.Cursor = Cursors.Hand;
            button.Font = new Font("Arial", 10F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
        }

        private void ApplyToolTips()
        {
            toolTip.SetToolTip(txtId, "Введіть числовий ID для реєстрації.");
            toolTip.SetToolTip(txtTitle, "Введіть назву твору.");
            toolTip.SetToolTip(txtAuthor, "Введіть ім'я автора.");
            toolTip.SetToolTip(dateTimePicker, "Виберіть дату реєстрації.");
            toolTip.SetToolTip(txtType, "Введіть тип твору (наприклад, книга, пісня).");
        }

        private void RefreshText()
        {
            richTextBox.Clear();
            richTextBox.AppendText($"{"ID",-5} {"Назва",-30} {"Автор",-20} {"Дата",-15} {"Тип",-10}\n");
            richTextBox.AppendText("--------------------------------------------------------------------------\n");
            foreach (var registration in manager.GetAllRegistrations())
            {
                richTextBox.AppendText($"{registration.Id,-5} {registration.Title,-30} {registration.Author,-20} {registration.RegistrationDate.ToShortDateString(),-15} {registration.Type,-10}\n");
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            int id;
            if (!int.TryParse(txtId.Text, out id))
            {
                MessageBox.Show("Некоректний формат ID!");
                return;
            }

            string title = txtTitle.Text;
            string author = txtAuthor.Text;
            DateTime regDate = dateTimePicker.Value;
            string type = txtType.Text;

            var newRegistration = new CopyrightRegistration(id, title, author, regDate, type);
            manager.AddRegistration(newRegistration);

            RefreshText();
            ClearInputFields();
        }

        private void ClearInputFields()
        {
            txtId.Clear();
            txtTitle.Clear();
            txtAuthor.Clear();
            txtType.Clear();
            dateTimePicker.Value = DateTime.Now;
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Document (*.doc)|*.doc";
                saveFileDialog.FilterIndex = 1;
                saveFileDialog.RestoreDirectory = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    try
                    {
                        var wordApp = new Word.Application();
                        var doc = wordApp.Documents.Add();
                        var table = doc.Tables.Add(doc.Range(), manager.GetAllRegistrations().Count + 1, 5); // Создание таблицы

                        // Заголовки столбцов
                        table.Cell(1, 1).Range.Text = "ID";
                        table.Cell(1, 2).Range.Text = "Назва";
                        table.Cell(1, 3).Range.Text = "Автор";
                        table.Cell(1, 4).Range.Text = "Дата";
                        table.Cell(1, 5).Range.Text = "Тип";

                        int row = 2;
                        foreach (var registration in manager.GetAllRegistrations())
                        {
                            table.Cell(row, 1).Range.Text = registration.Id.ToString();
                            table.Cell(row, 2).Range.Text = registration.Title;
                            table.Cell(row, 3).Range.Text = registration.Author;
                            table.Cell(row, 4).Range.Text = registration.RegistrationDate.ToShortDateString();
                            table.Cell(row, 5).Range.Text = registration.Type;
                            row++;
                        }

                        doc.SaveAs2(filePath);
                        doc.Close();
                        wordApp.Quit();

                        MessageBox.Show("Дані успішно експортовано у файл!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Помилка при збереженні файлу: {ex.Message}");
                    }
                }
            }
        }




        private void ParseAndLoadData(string text)
        {
            richTextBox.Clear();
            richTextBox.AppendText($"{"ID",-5} {"Назва",-30} {"Автор",-20} {"Дата",-15} {"Тип",-10}\n");
            richTextBox.AppendText("--------------------------------------------------------------------------\n");

            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                string[] fields = line.Split(':');
                if (fields.Length >= 5)
                {
                    int id;
                    if (!int.TryParse(fields[1].Trim(), out id))
                        continue;

                    string title = fields[2].Trim();
                    string author = fields[3].Trim();
                    DateTime regDate;
                    if (!DateTime.TryParse(fields[4].Trim(), out regDate))
                        continue;

                    string type = fields[5].Trim(); // Было изменено на 5

                    richTextBox.AppendText($"{id,-5} {title,-30} {author,-20} {regDate.ToShortDateString(),-15} {type,-10}\n");
                }
            }
        }


        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            RefreshText();
        }

       
    }
}



