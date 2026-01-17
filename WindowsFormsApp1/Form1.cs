using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string templatePath = @"C:\Users\svist\Desktop\институт\3 курс\5 семестр\КРПП\WindowsFormsApp1\template.docx";
            string savePath = @"C:\Users\svist\Desktop\институт\3 курс\5 семестр\КРПП\WindowsFormsApp1\Gotoviy_Akt.docx";

            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(templatePath);

                ReplaceStub("{Метка1}", textBox1.Text, doc);
                ReplaceStub("{Метка2}", textBox2.Text, doc);
                ReplaceStub("{Метка3}", textBox3.Text, doc);
                ReplaceStub("{Метка4}", textBox4.Text, doc);
                ReplaceStub("{Метка5}", textBox5.Text, doc);
                ReplaceStub("{Метка6}", textBox6.Text, doc);
                ReplaceStub("{Метка7}", textBox7.Text, doc);
                ReplaceStub("{Метка8}", textBox8.Text, doc);
                ReplaceStub("{Метка9}", textBox9.Text, doc);
                ReplaceStub("{Метка10}", textBox10.Text, doc);

                doc.SaveAs2(savePath);

                MessageBox.Show("Документ успешно создан по адресу: " + savePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
            finally
            {
                if (doc != null) doc.Close();
                wordApp.Quit();
            }

        }
        private void ReplaceStub(string stub, string text, Word.Document doc)
        {
            var range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stub, ReplaceWith: text);
        }
    }
}
