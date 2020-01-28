using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.ObjectModel;

//using Microsoft.Office.Interop.Word; to edit and write word documents

namespace coverLetterGenerator
{
    public partial class CLG : Form
    {
        public CLG()
        {
            InitializeComponent();
        }

        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            //required variable for usind the find and replace method of the libirary
            object matchCase = true;
            object matchWholeWord = true;
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

            bool replaced = wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
            Console.WriteLine(replaced);

        }

        //Creeate the Doc Method
        private void CreateWordDocument(object location)
        {
            // repeat the replacing and saving for each row in the grid
            foreach (DataGridViewRow row in DGVData.Rows)
            {
                if (row.Index == DGVData.RowCount-1) { break; }
                Word.Application wordApp = new Word.Application();
                object missing = Missing.Value;
                Word.Document myWordDoc = null;
                // use CoverTemp.docx as template the location is the path of exe file on the debug or release folder
                object filename = ((string)location) + "CoverTemp.docx";
                // name of the file that will be generated
                object SaveAs= ((string)location) + "Moussab_MOULIM_CoverLetter_"+ row.Cells[0].Value.ToString().Replace(" ","_")+ ".docx";

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
                    this.FindAndReplace(wordApp, "<recipient>", row.Cells[2].Value.ToString());
                    this.FindAndReplace(wordApp, "<position>", row.Cells[1].Value.ToString());
                    this.FindAndReplace(wordApp, "<company>", row.Cells[0].Value.ToString());
                    this.FindAndReplace(wordApp, "<finish>", row.Cells[3].Value.ToString().Equals("y")?"Your sincerely":"Your faithfully");
                }
                else
                {
                    MessageBox.Show("File not Found!. Please make sure the template is in the same directory as the program.");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                

                
            }
            MessageBox.Show("Cover Letters Created!");

        }


        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            //the path of the exe file generally on the bin/debug folder or pn bin/release folder base on your mode of compiling the app
            CreateWordDocument(Application.StartupPath + @"\");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Navigate to a URL.
            System.Diagnostics.Process.Start("http://moussaab-moulim.github.io/");
        }
    }
}
