using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ExcelFunctionLibrary;
using ExcelFunctionLibrary.Models;
using TRVautomator;
using System.Globalization;
using ExcelFunctionLibrary.Models.RWGautomator;
using Microsoft.Office.Interop.Excel;

namespace KRGautomator
{
    public partial class RWGautomator : Form
    {
        private OpenFileDialog openFileDialog = new();
        Excel.Application? xlApp = null;
        Excel.Workbook? xlWorkbook = null;
        Excel.Worksheet? xlWorksheet = null;
        private bool IsTraining = false;
        private bool IsAdmin = false;
        public RWGautomator()
        {
            LoginForm form = new();
            form.ShowDialog();
            if (form.showAutomator)
            {
                InitializeComponent();

                IsTraining = form.IsTraining;
                IsAdmin = form.IsAdmin;
                if (IsTraining)
                {
                    this.Text += " (Training)";
                }
            }
            else
            {
                Environment.Exit(0);
            }

        }

        private void generateRoots_Click(object sender, EventArgs e)
        {
            try
            {

                string fileTARGET = System.IO.Path.GetFullPath(openFileDialog.FileName);

                if (fileTARGET == null)
                {
                    MessageBox.Show("Please select a valid document.");
                }
                else
                {

                    generateRoots.BackColor = Color.Red;
                    xlWorksheet = xlWorkbook.Sheets["Root Word Generator"];

                    List<DefinedColumn> readableColumns = new();
                    readableColumns.Add(new DefinedColumn() { column = (int)ExcelColumns.B, columnStartsFrom = 3 });
                    readableColumns.Add(new DefinedColumn() { column = (int)ExcelColumns.C, columnStartsFrom = 3 });
                    ReadBasicColumnData readExcel = new(xlWorksheet, readableColumns);
                    WriteBasicColumnData writeExcel = new(xlWorksheet);
                    
                    generateRoots.Text = "Generating unique roots...";
                    List<StackedColumns> primaryKeywords = readExcel.ReadStackedColumns();
                    List<RWGWritableData> uniqueWords = new();
                    int highestWordCount = 0;
                    foreach (StackedColumns sentence in primaryKeywords)
                    {
                        highestWordCount = sentence.stackedCellValues[0].Split(" ").Length > highestWordCount ? sentence.stackedCellValues[0].Split(" ").Length : highestWordCount;

                        int searchVolume = int.Parse(sentence.stackedCellValues[1], NumberStyles.AllowThousands);
                        foreach (var word in sentence.stackedCellValues[0].Split(" ").Select((value, i) => new {i, value}))
                        {
                            try
                            {
                                List<RWGWritableData> isUnique = uniqueWords.Where(s1 => s1.keyword.Contains(word.value)).ToList();
                                if (isUnique.Count == 0)
                                {
                                    uniqueWords.Add(new RWGWritableData() { keyword = word.value, frequency = 1, searchVolume = searchVolume });
                                }
                                else
                                {
                                    foreach (RWGWritableData nonUnique in isUnique)
                                    {
                                        if (nonUnique.keyword == word.value)
                                        {
                                            nonUnique.frequency++;
                                            nonUnique.searchVolume += searchVolume;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch { uniqueWords.Add(new RWGWritableData() { keyword = word.value, frequency = 1, searchVolume = searchVolume }); }

                        }
                    }


                    generateRoots.Text = "Generating unique broad roots of 2...";
                    List<RWGWritableData> unique2WordsBroad = GenerateRoots.ExpandSentencesBroad(uniqueWords, uniqueWords);
                    unique2WordsBroad = GenerateRoots.GenerateWordCountBroad(unique2WordsBroad, primaryKeywords);
                    
                    generateRoots.Text = "Generating unique exact roots of 2...";
                    List<RWGWritableData> unique2WordsExact = GenerateRoots.ExpandSentencesExact(uniqueWords, uniqueWords);
                    unique2WordsExact = GenerateRoots.GenerateWordCountExact(unique2WordsExact, primaryKeywords);




                    if (highestWordCount > 2)
                    {
                        generateRoots.Text = "Generating unique broad roots of 3...";
                        List<RWGWritableData> unique3WordsBroad = GenerateRoots.ExpandSentencesBroad(unique2WordsBroad, uniqueWords);
                        unique3WordsBroad = GenerateRoots.GenerateWordCountBroad(unique3WordsBroad, primaryKeywords);

                        generateRoots.Text = "Generating unique exact roots of 3...";
                        List<RWGWritableData> unique3WordsExact = GenerateRoots.ExpandSentencesExact(unique2WordsExact, uniqueWords);
                        unique3WordsExact = GenerateRoots.GenerateWordCountExact(unique3WordsExact, primaryKeywords);


                        unique3WordsBroad = unique3WordsBroad.Where(s1 => s1.frequency > 0).ToList();
                        unique3WordsBroad.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));
                        unique3WordsExact = unique3WordsExact.Where(s1 => s1.frequency > 0).ToList();
                        unique3WordsExact.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));

                        writeExcel.WriteRWGCustomData((int)ExcelColumns.Q, (int)ExcelColumns.R, (int)ExcelColumns.S, unique3WordsBroad, 3);
                        writeExcel.WriteRWGCustomData((int)ExcelColumns.U, (int)ExcelColumns.V, (int)ExcelColumns.W, unique3WordsExact, 3);


                        if (highestWordCount > 3)
                        {
                            generateRoots.Text = "Generating unique broad roots of 4...";
                            List<RWGWritableData> unique4WordsBroad = GenerateRoots.ExpandSentencesBroad(unique3WordsBroad, uniqueWords);
                            unique4WordsBroad = GenerateRoots.GenerateWordCountBroad(unique4WordsBroad, primaryKeywords);

                            generateRoots.Text = "Generating unique exact roots of 4...";
                            List<RWGWritableData> unique4WordsExact = GenerateRoots.ExpandSentencesExact(unique3WordsExact, uniqueWords);
                            unique4WordsExact = GenerateRoots.GenerateWordCountExact(unique4WordsExact, primaryKeywords);


                            unique4WordsBroad = unique4WordsBroad.Where(s1 => s1.frequency > 0).ToList();
                            unique4WordsBroad.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));
                            unique4WordsExact = unique4WordsExact.Where(s1 => s1.frequency > 0).ToList();
                            unique4WordsExact.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));

                            writeExcel.WriteRWGCustomData((int)ExcelColumns.Y, (int)ExcelColumns.Z, (int)ExcelColumns.AA, unique4WordsBroad, 3);
                            writeExcel.WriteRWGCustomData((int)ExcelColumns.AC, (int)ExcelColumns.AD, (int)ExcelColumns.AE, unique4WordsExact, 3);


                            if (highestWordCount > 4)
                            {
                                generateRoots.Text = "Generating unique broad roots of 5...";
                                List<RWGWritableData> unique5WordsBroad = GenerateRoots.ExpandSentencesBroad(unique4WordsBroad, uniqueWords);
                                unique5WordsBroad = GenerateRoots.GenerateWordCountBroad(unique5WordsBroad, primaryKeywords);

                                generateRoots.Text = "Generating unique exact roots of 5...";
                                List<RWGWritableData> unique5WordsExact = GenerateRoots.ExpandSentencesExact(unique4WordsExact, uniqueWords);
                                unique5WordsExact = GenerateRoots.GenerateWordCountExact(unique5WordsExact, primaryKeywords);


                                unique5WordsBroad = unique5WordsBroad.Where(s1 => s1.frequency > 0).ToList();
                                unique5WordsBroad.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));
                                unique5WordsExact = unique5WordsExact.Where(s1 => s1.frequency > 0).ToList();
                                unique5WordsExact.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));

                                writeExcel.WriteRWGCustomData((int)ExcelColumns.AG, (int)ExcelColumns.AH, (int)ExcelColumns.AI, unique5WordsBroad, 3);
                                writeExcel.WriteRWGCustomData((int)ExcelColumns.AK, (int)ExcelColumns.AL, (int)ExcelColumns.AM, unique5WordsExact, 3);


                            }
                        }

                    }


                    uniqueWords = uniqueWords.Where(s1 => s1.frequency > 0).ToList();
                    uniqueWords.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));
                    unique2WordsBroad = unique2WordsBroad.Where(s1 => s1.frequency > 0).ToList();
                    unique2WordsBroad.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));
                    unique2WordsExact = unique2WordsExact.Where(s1 => s1.frequency > 0).ToList();
                    unique2WordsExact.Sort((s1, s2) => s2.frequency.CompareTo(s1.frequency));



                    generateRoots.Text = "Writing roots...";

                    writeExcel.WriteRWGCustomData((int)ExcelColumns.E, (int)ExcelColumns.F, (int)ExcelColumns.G, uniqueWords, 3);
                    writeExcel.WriteRWGCustomData((int)ExcelColumns.I, (int)ExcelColumns.J, (int)ExcelColumns.K, unique2WordsBroad, 3);
                    writeExcel.WriteRWGCustomData((int)ExcelColumns.M, (int)ExcelColumns.N, (int)ExcelColumns.O, unique2WordsExact, 3);


                    xlWorkbook.Save();
                    generateRoots.BackColor = Color.White;
                    generateRoots.Text = "Generate roots";
                }
            }
            catch (Exception ex)
            {
                generateRoots.BackColor = Color.White;
                generateRoots.Text = "Error occured - check message box";
                MessageBox.Show(ex.Message);
            }
        }

        private void uploadDocument_Click(object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = "C://Desktop";
            openFileDialog.Title = "Select file to be upload.";
            openFileDialog.Filter = "Select Valid Document(*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 1;
            if (xlWorkbook != null)
            {
                try
                {
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);
                    xlWorkbook = null;
                }
                catch { }
            }
            if (xlApp != null)
            {
                try
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                    xlApp = null;
                }
                catch { }

            }
            try
            {
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(openFileDialog.FileName);
                        uploadDocument.BackColor = Color.Red;
                        xlApp = new();
                        xlWorkbook = xlApp.Workbooks.Open(path);
                        xlApp.DisplayAlerts = false;
                        xlApp.Visible = true;
                    }
                }
                else
                {
                    MessageBox.Show("Please Upload KRG workbook.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void KRGautomator_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (xlWorksheet != null) Marshal.FinalReleaseComObject(xlWorksheet);
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.FinalReleaseComObject(xlWorkbook);
            }
            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
            }
        }
    }
}