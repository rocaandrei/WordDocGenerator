using System;
using System.Drawing;
using MetroFramework.Forms;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Windows.Forms;

namespace WordDocApp
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Metoda se uita in fisierul nostru word si inlocuieste textul introdus dupa cerintele noastre
        /// Este metoda Replace din documentul word (buton, colt dreapta sus)
        /// </summary>
        private void FindAndReplaceWords(Microsoft.Office.Interop.Word.Application wordApp, object toFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildcards = false;
            object matchSoundsLike = false;
            object MatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacrits = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object readOnly = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            //documentatie: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.find.execute?view=word-pia
            wordApp.Selection.Find.Execute(ref toFindText, ref matchCase, ref matchWholeWord,
                ref matchWildcards, ref matchSoundsLike, ref MatchAllWordForms, ref forward, ref wrap, ref format,
             ref replaceWithText, ref replace, ref matchKashida, ref matchDiacrits, ref matchAlefHamza, ref matchControl);
        }

        private void GenerateWordDocument(object fileName, object saveAsFileName)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            //2. ne generam un obiect Missing -> ii spunem metodei de jos ca nu vrem sa folosim acel parametru
            object missing = Missing.Value;
            //3. acesta este documentul ce va fi generat si in care ne vom salva informatia
            Document wordFile = new Microsoft.Office.Interop.Word.Document();

            //4. verificam daca exista acest fisier
            bool isValidFile = File.Exists((string)fileName);
            if (isValidFile)
            {
                //5. readonliy e de fapt true, doar ca la deschidere trebuie sa fie false pentru a putea modifica documentul
                object readOnly = false;
                //6 isVisibile = false, pentru ca nu ne interesaza bucataria interna si ce se intampla acolo
                object isVisible = false;
                //7. wordApp.Visibile = false pentru ca nu vrem sa ne apara documentul la generare pe ecran, daca vrem asta punem true.
                wordApp.Visible = false;

                //8. aici deschidem fisierul Word, dupa care ii setam <fileName>, trecem parametri missing 
                // pentru ca nu sunt necesari si dupa care setam optiunea de readOnly pentru a putea modifica documentul
                wordFile = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                wordFile.Activate();

                //9. aici facem modificarea cuvintelor
                FindAndReplaceWords(wordApp, "<name>", nameTextBox.Text);
                FindAndReplaceWords(wordApp, "<firstName>", firstNameTextBox.Text);
                FindAndReplaceWords(wordApp, "<date>", DateTime.Now.ToString("dd.MM.yyyy"));
            }
            else
            {
                //daca nu functioneaza apare mesajul
                System.Windows.Forms.MessageBox.Show("File not found!", "File not found!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            //10. Salvam documentul, saveAs va fi calea catre documentul nostru
            wordFile.SaveAs2( ref saveAsFileName, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //11. inchidem fisierul Word
            wordFile.Close();
            //12. inchidem aplicatia Word
            wordApp.Quit();

            //mesaj de feedback- ca a functionat
            MessageBox.Show("File done!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void generateDocButton_MouseHover(object sender, EventArgs e)
        {
            generateDocButton.BackColor = Color.Teal;
        }

        private void generateDocButton_MouseLeave(object sender, EventArgs e)
        {
            generateDocButton.BackColor = Color.White;
        }

        private void generateDocButton_Click(object sender, EventArgs e)
        {
            object pathToFileName = @"C:\Users\AndreiR\Documents\Visual Studio 2015\Projects\WordGenerator\WordApp\WordDocApp\Certificate_template.docx";
            object pathToSaveAs = @"C:\Users\AndreiR\Documents\Visual Studio 2015\Projects\WordGenerator\WordApp\WordDocApp\Certificate_template1.docx";
            GenerateWordDocument(pathToFileName, pathToSaveAs);
        }
    }
}
