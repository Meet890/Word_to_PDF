using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Text.RegularExpressions;
using System.Reflection.Emit;


namespace Epic
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)

        {
            object matchCase = true;
            object matchWholeword = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
            ref matchCase, ref matchWholeword,
            ref matchWildCards, ref matchSoundLike,
            ref nmatchAllforms, ref forward,
            ref wrap, ref format, ref replaceWithText,
            ref replace, ref matchKashida, ref matchDiactitics,
            ref matchAlefHamza, ref matchControl);
        }


        private void CreateWordDocument(object filename, object SaveAs)

        {

            Word.Application wordApp = new Word.Application();

            object missing = Missing.Value;

            Word.Document myWordDoc = null;
            string a = (string)filename + (string)SaveAs;

            MessageBox.Show(a);
            if (File.Exists((string)filename))

            {

                object readOnly = false;

                object isVisible = false;

                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,

                ref missing, ref missing, ref missing,

                ref missing, ref missing, ref missing,

                ref missing, ref missing, ref missing,

                ref missing, ref missing, ref missing, ref missing);

                myWordDoc.Activate();

                //find and replace

                this.FindAndReplace(wordApp, label1.Text, textBox1.Text);

                this.FindAndReplace(wordApp, label2.Text, textBox2.Text);
                this.FindAndReplace(wordApp, label3.Text, textBox3.Text);
                this.FindAndReplace(wordApp, label4.Text, textBox4.Text);

                //this.FindAndReplace(wordApp, "<birthday>", dateTimePicker1.Value.ToShortDateString());

                //this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToShortDateString());

            }

            else

            {



                MessageBox.Show("File not Found!");
            }



            //Save as
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            myWordDoc.SaveAs(ref SaveAs, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File Created!");
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreateWordDocument(@"C:\Users\Marble\Desktop\marble.docx", @"C:\Users\Marble\Desktop\temp.docx");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc",
                Title = "Select a Word Document"
            };
            
            //try
            //{
            //    // Open File Dialog to select a Word document


            //    if (openFileDialog.ShowDialog() == DialogResult.OK)
            //    {
            //        string filePath = openFileDialog.FileName;

            //        // Create Word application and document objects
            //        Word.Application wordApp = new Word.Application();
            //        Document wordDoc = null;

            //        try
            //        {
            //            wordDoc = wordApp.Documents.Open(filePath);

            //            // List to store red text
            //            List<string> redTextList = new List<string>();
            //            //string[] red = new string[];
            //            // Loop through all ranges in the document's content
            //            foreach (Word.Range word in wordDoc.Words)
            //            {
            //                // Check if the range text is not whitespace and is red
            //                if (!string.IsNullOrWhiteSpace(word.Text))
            //                {
            //                    // Use Regex to check if the text matches the <text> pattern
            //                    if (Regex.IsMatch(word.Text.Trim(), @"^<.*>$"))
            //                    {
            //                        // Add the matching text to the list
            //                        redTextList.Add(word.Text.Trim());

            //                    }
            //                }
            //            }

            //            // Display all red text
            //            //if (redTextList.Count > 0)
            //            //{
            //            //    string redTextDisplay = string.Join(Environment.NewLine, redTextList);
            //            //    MessageBox.Show($"Red Text Found:\n\n{redTextDisplay}", "Red Text");
            //            //}
            //            //else
            //            //{
            //            //    MessageBox.Show("No red text found in the document.", "Result");
            //            //}
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show($"Error while processing the document: {ex.Message}", "Error");
            //        }
            //        finally
            //        {
            //            // Properly close and release Word resources
            //            if (wordDoc != null)
            //                wordDoc.Close(false);

            //            wordApp.Quit();
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"An error occurred: {ex.Message}", "Error");
            //}



            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                string redText = ExtractRedText(filePath);
                //string[] redText = tempText.ToArray();
                richTextBox1.Text = redText;
            }
        }
        private string ExtractRedText(string filePath)
        {
            List<string> redTextList = new List<string>();
            Word.Application wordApp = new Word.Application();
            Document wordDoc = null;
            string redText = string.Empty;
            List<string> colorsList = new List<string> { };
            //string[] redArray = colorsList.ToArray();
            try
            {
                wordDoc = wordApp.Documents.Open(filePath);

                //foreach (Word.Range range in wordDoc.StoryRanges)
                //{
                //    foreach (Microsoft.Office.Interop.Word.Range word in range.Words)
                //    {
                //        if (word.Font.Color == WdColor.wdColorRed)
                //        {
                //            redText += word.Text;
                //            colorsList.Add(word.Text);
                //        }
                //    }
                //}

                // List to store red text
                //List<string> redTextList = new List<string>();
                ////string[] red = new string[];
                //// Loop through all ranges in the document's content
                //foreach (Word.Range word in wordDoc.Words)
                //{
                //    // Check if the range text is not whitespace and is red
                //    if (!string.IsNullOrWhiteSpace(word.Text))
                //    {
                //        // Use Regex to check if the text matches the <text> pattern
                //        if (Regex.IsMatch(word.Text.Trim(), @"^<.*>$"))
                //        {
                //            // Add the matching text to the list
                //            //redTextList.Add(word.Text.Trim());
                //            colorsList.Add(word.Text.Trim());

                //        }
                //    }
                //}
                var a = 0;
                string text = "";
                foreach (Microsoft.Office.Interop.Word.Range word in wordDoc.Words)
                {

                    var match = Regex.Match(word.Text, @"<");
                    if (match.Success)
                    {
                        a = 1;
                    }
                    var match2 = Regex.Match(word.Text, @">");
                    if (match2.Success)
                    {
                        string textInsideBrackets = word.Text;
                        text = text + textInsideBrackets;
                        redTextList.Add(text);
                        a = 0;
                        text = "";
                    }
                    if (a == 1)
                    {
                        //var match2 = Regex.Match(word.Text, @">");
                        // The text between the angle brackets is captured in Group 1
                        string textInsideBrackets = word.Text;
                        text = text + textInsideBrackets;

                    }
                    else
                    {
                        //hello
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Clean up
                wordDoc?.Close(false);
                wordApp.Quit();
            }
            string[] redArray = redTextList.ToArray();

            for (int i = 0; i < redArray.Length; i++)
            {
                // Dynamically find the label by its name
                Control label = this.Controls.Find($"label{i + 1}", true).FirstOrDefault();

                if (label != null)
                {
                    // Set the text of the label
                    label.Text = redArray[i];
                }


            }
            // Create a new Label




            return redText;
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }
    }
}
