using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace pdfsm {
    public partial class Form1 : Form {

        private string ConfigFile = ".\\pdfsm_config.ini";
        private string SelectFile = "";
        private string InputPath = "";
        private string SaveFile = "";
        private bool ErrorsOccurred = false;

        public Form1() {
            InitializeComponent();
        }

        #region Functions
        /// <summary>Checks if the merge button is available (all input fields have to be filled).</summary>
        private void MergeButtonAvailable() {
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals("") && !textBox3.Text.Equals("")) {
                button4.Enabled = true;
            } else {
                button4.Enabled = false;
            }
        }

        /// <summary>Writes entered paths to a config file to remember them next time (if possible).</summary>
        private void Write(string key, string value) {
            try {
                IniFileHelper.WriteValue("General", key, value, ConfigFile);
            } catch (Exception) {
                Append("Tip: If you move this program to a writable directory, it could remember your configured paths. ;)\n");
            }
        }

        /// <summary>Merges the PDF Documents based on the INI Selector File.</summary>
        private void Merge() {
            button6.Enabled = false; richTextBox1.Clear(); ErrorsOccurred = false;

            // Preparation
            string Description = Read("General", "Description");
            int InputFilesCount = Convert.ToInt32(Read("General", "InputFilesCount", "0"));
            string InputFileName = ""; string SelectPages = ""; string[] SelectPagesArray;
            PdfDocument PdfIn = null; PdfDocument PdfOut = new PdfDocument();

            if (Description != "") {
                Append(Description + "\n\n");
            }

            // For each file
            for (int i = 1; i <= InputFilesCount; i++) {
                // Preparation
                Append("Using file " + i + " of " + InputFilesCount + "...\n");
                InputFileName = Read("Input" + i, "InputFile");

                // Check existance
                if (File.Exists(InputPath + InputFileName + ".pdf")) {
                    try {
                        PdfIn = PdfReader.Open(InputPath + InputFileName + ".pdf", PdfDocumentOpenMode.Import);
                    } catch (Exception) {
                        Append("Fixing corrupted file " + InputFileName + ".pdf...\n");
                        Process proc = Process.Start(".\\gswin32c.exe", "-o \"" + Path.GetDirectoryName(SaveFile) + "\\pdfsm_fixed.pdf\" -sDEVICE=pdfwrite -dPDFSETTINGS=/prepress \"" + InputPath + InputFileName + ".pdf\"");
                        proc.WaitForExit();
                        PdfIn = PdfReader.Open(Path.GetDirectoryName(SaveFile) + "\\pdfsm_fixed.pdf", PdfDocumentOpenMode.Import);
                    }
                } else {
                    Append("Skipping file because not found in Input Path:\n" + InputPath + InputFileName + ".pdf\n\n");
                    continue;
                }

                // Read pages
                SelectPages = Read("Input" + i, "SelectPages");
                SelectPagesArray = SelectPages.Split(',');

                // For each page in file
                foreach (string p in SelectPagesArray) {
                    Append("Copying page " + p + "...\n");
                    CopyPage(PdfIn, PdfOut, Convert.ToInt32(p));  // Actual copy process
                }

                Append("\n");
            }

            // Finish
            if (File.Exists(Path.GetDirectoryName(SaveFile) + "\\pdfsm_fixed.pdf")) {
                File.Delete(Path.GetDirectoryName(SaveFile) + "\\pdfsm_fixed.pdf");
            }

            if (PdfOut.PageCount > 0) {
                try {
                    PdfOut.Save(SaveFile);  // Actual save process
                    button6.Enabled = true;

                    if (!ErrorsOccurred) {
                        Append("Finished.");
                    } else {
                        Append("Finished with errors (check the log!).");
                    }
                } catch (Exception ex) {
                    Append("Could not save output file:\n" + ex.Message);
                }
            } else {
                Append("No output file has been saved as it would be empty (either input files missing or selection file wrong).");
            }
        }

        /// <summary>Copies a page from a PdfDocument to a PdfDocument.</summary>
        private void CopyPage(PdfDocument from, PdfDocument to, int pageNumber) {
            try {
                to.AddPage(from.Pages[pageNumber - 1]);
            } catch (Exception ex) {
                Append("Skipping page " + pageNumber + " (does it exist?):\n" + ex.Message + "\n");
                ErrorsOccurred = true;
            }
        }

        /// <summary>Reads a key from an INI file and returns its value as a string.</summary>
        public string Read(string section, string key, string defaultValue = "", string file = "") {
            if (file.Equals("")) { file = SelectFile; };
            return IniFileHelper.ReadValue(section, key, file, defaultValue);
        }

        /// <summary>Appends text to the log.</summary>
        private void Append(String text) {
            richTextBox1.AppendText(text);
        }
        #endregion

        #region Button Events
        private void button1_Click(object sender, EventArgs e) {
            string c = Read("General", "SelectFile", "", ConfigFile);
            if (!c.Equals("")) {
                openFileDialog1.FileName = Path.GetFileName(c);
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(c);
            }

            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                SelectFile = openFileDialog1.FileName;
                textBox1.Text = SelectFile;

                saveFileDialog1.FileName = IniFileHelper.ReadValue("General", "DefaultOutputName", SelectFile, "");
                MergeButtonAvailable();
                Write("SelectFile", SelectFile);
            }
        }

        private void button2_Click(object sender, EventArgs e) {
            string c = Read("General", "InputPath", "", ConfigFile);
            if (!c.Equals("")) {
                vistaFolderBrowserDialog1.SelectedPath = c;
            }

            if (vistaFolderBrowserDialog1.ShowDialog() == DialogResult.OK) {
                InputPath = vistaFolderBrowserDialog1.SelectedPath + "\\";
                textBox2.Text = InputPath;
                MergeButtonAvailable();
                Write("InputPath", InputPath);
            }
        }

        private void button3_Click(object sender, EventArgs e) {
            string c = Read("General", "SaveFile", "", ConfigFile);
            if (!c.Equals("")) {
                saveFileDialog1.FileName = Path.GetFileName(c);
                saveFileDialog1.InitialDirectory = Path.GetDirectoryName(c);

                string k = Read("General", "DefaultOutputName");
                if (!k.Equals("")) {
                    saveFileDialog1.FileName = k;
                }
            }

            if (saveFileDialog1.ShowDialog() == DialogResult.OK) {
                SaveFile = saveFileDialog1.FileName;
                textBox3.Text = SaveFile;
                MergeButtonAvailable();
                Write("SaveFile", SaveFile);
            }
        }

        private void button4_Click(object sender, EventArgs e) {
            Merge();
        }

        private void button5_Click(object sender, EventArgs e) {
            MessageBox.Show(this, "This program selects pages from a set of PDF files specified in an INI Selector File and merges them to an output file.\n\nIt includes PDFsharp to copy pages and Ghostscript to fix corrupted files.", "PDF Select & Merge", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button6_Click(object sender, EventArgs e) {
            if (File.Exists(SaveFile)) {
                Process.Start(SaveFile);
            } else {
                MessageBox.Show(this, "File has been removed.", "PDF Select & Merge", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Process.Start("https://github.com/woborschilde/pdfsm/releases");
        }
        #endregion

    }
}
