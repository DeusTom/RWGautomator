using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ExcelFunctionLibrary;

namespace KRGautomator
{
    public partial class RWGautomator : Form
    {
        private OpenFileDialog openFileDialog = new();
        Excel.Application? xlApp = null;
        Excel.Workbook? xlWorkbook = null;
        Excel.Worksheet? xlWorksheet = null;
        public RWGautomator()
        {
            InitializeComponent();
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

                    xlWorksheet = xlWorkbook.Sheets["ROOT WORD GENERATOR"];
                    List<DefinedColumn> readableColumns = new();
                    readableColumns.Add(new DefinedColumn() { column = (int)ExcelColumns.B, columnStartsFrom = 3 });
                    ReadBasicColumnData readExcel = new(xlWorksheet, readableColumns);
                    WriteBasicColumnData writeExcel = new(xlWorksheet);

                    generateRoots.Text = "Generating unique roots...";
                    List<string> primaryKeywords = readExcel.ReadColumns();
                    List<ExcelFunctionLibrary.StringAndInt> uniqueWords = new();
                    int highestWordCount = 0;
                    foreach (string word in primaryKeywords)
                    {
                        highestWordCount = word.Split(" ").Length > highestWordCount ? word.Split(" ").Length : highestWordCount;

                        foreach (string splittedWord in word.Split(" "))
                        {
                            try
                            {
                                List<ExcelFunctionLibrary.StringAndInt> isUnique = uniqueWords.Where(s1 => s1.keyword.Contains(splittedWord)).ToList();
                                if (isUnique.Count == 0)
                                {
                                    uniqueWords.Add(new ExcelFunctionLibrary.StringAndInt() { keyword = splittedWord, number = 1 });
                                }
                                else
                                {
                                    foreach (ExcelFunctionLibrary.StringAndInt nonUnique in isUnique)
                                    {
                                        if (nonUnique.keyword == splittedWord)
                                        {
                                            nonUnique.number++;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch { uniqueWords.Add(new ExcelFunctionLibrary.StringAndInt() { keyword = splittedWord, number = 1 }); }

                        }
                    }


                    generateRoots.Text = "Generating unique broad roots of 2...";
                    List<ExcelFunctionLibrary.StringAndInt> unique2WordsBroad = GenerateRoots.ExpandSentencesBroad(uniqueWords, uniqueWords);
                    unique2WordsBroad = GenerateRoots.GenerateWordCountBroad(unique2WordsBroad, primaryKeywords);
                    
                    generateRoots.Text = "Generating unique exact roots of 2...";
                    List<ExcelFunctionLibrary.StringAndInt> unique2WordsExact = GenerateRoots.ExpandSentencesExact(uniqueWords, uniqueWords);
                    unique2WordsExact = GenerateRoots.GenerateWordCountExact(unique2WordsExact, primaryKeywords);




                    if (highestWordCount > 2)
                    {
                        generateRoots.Text = "Generating unique broad roots of 3...";
                        List<ExcelFunctionLibrary.StringAndInt> unique3WordsBroad = GenerateRoots.ExpandSentencesBroad(unique2WordsBroad, uniqueWords);
                        unique3WordsBroad = GenerateRoots.GenerateWordCountBroad(unique3WordsBroad, primaryKeywords);

                        generateRoots.Text = "Generating unique exact roots of 3...";
                        List<ExcelFunctionLibrary.StringAndInt> unique3WordsExact = GenerateRoots.ExpandSentencesExact(unique2WordsExact, uniqueWords);
                        unique3WordsExact = GenerateRoots.GenerateWordCountExact(unique3WordsExact, primaryKeywords);


                        unique3WordsBroad = unique3WordsBroad.Where(s1 => s1.number > 0).ToList();
                        unique3WordsBroad.Sort((s1, s2) => s2.number.CompareTo(s1.number));
                        unique3WordsExact = unique3WordsExact.Where(s1 => s1.number > 0).ToList();
                        unique3WordsExact.Sort((s1, s2) => s2.number.CompareTo(s1.number));

                        writeExcel.WriteStringAndInt((int)ExcelColumns.N, (int)ExcelColumns.M, unique3WordsBroad, 3);
                        writeExcel.WriteStringAndInt((int)ExcelColumns.Q, (int)ExcelColumns.P, unique3WordsExact, 3);

                        if (highestWordCount > 3)
                        {
                            generateRoots.Text = "Generating unique broad roots of 4...";
                            List<ExcelFunctionLibrary.StringAndInt> unique4WordsBroad = GenerateRoots.ExpandSentencesBroad(unique3WordsBroad, uniqueWords);
                            unique4WordsBroad = GenerateRoots.GenerateWordCountBroad(unique4WordsBroad, primaryKeywords);

                            generateRoots.Text = "Generating unique exact roots of 4...";
                            List<ExcelFunctionLibrary.StringAndInt> unique4WordsExact = GenerateRoots.ExpandSentencesExact(unique3WordsExact, uniqueWords);
                            unique4WordsExact = GenerateRoots.GenerateWordCountExact(unique4WordsExact, primaryKeywords);


                            unique4WordsBroad = unique4WordsBroad.Where(s1 => s1.number > 0).ToList();
                            unique4WordsBroad.Sort((s1, s2) => s2.number.CompareTo(s1.number));
                            unique4WordsExact = unique4WordsExact.Where(s1 => s1.number > 0).ToList();
                            unique4WordsExact.Sort((s1, s2) => s2.number.CompareTo(s1.number));    

                            writeExcel.WriteStringAndInt((int)ExcelColumns.T, (int)ExcelColumns.S, unique4WordsBroad, 3);
                            writeExcel.WriteStringAndInt((int)ExcelColumns.W, (int)ExcelColumns.V, unique4WordsExact, 3);

                            if (highestWordCount > 4)
                            {
                                generateRoots.Text = "Generating unique broad roots of 5...";
                                List<ExcelFunctionLibrary.StringAndInt> unique5WordsBroad = GenerateRoots.ExpandSentencesBroad(unique4WordsBroad, uniqueWords);
                                unique5WordsBroad = GenerateRoots.GenerateWordCountBroad(unique5WordsBroad, primaryKeywords);

                                generateRoots.Text = "Generating unique exact roots of 5...";
                                List<ExcelFunctionLibrary.StringAndInt> unique5WordsExact = GenerateRoots.ExpandSentencesExact(unique4WordsExact, uniqueWords);
                                unique5WordsExact = GenerateRoots.GenerateWordCountExact(unique5WordsExact, primaryKeywords);


                                unique5WordsBroad = unique5WordsBroad.Where(s1 => s1.number > 0).ToList();
                                unique5WordsBroad.Sort((s1, s2) => s2.number.CompareTo(s1.number));
                                unique5WordsExact = unique5WordsExact.Where(s1 => s1.number > 0).ToList();
                                unique5WordsExact.Sort((s1, s2) => s2.number.CompareTo(s1.number));

                                writeExcel.WriteStringAndInt((int)ExcelColumns.Z, (int)ExcelColumns.Y, unique5WordsBroad, 3);
                                writeExcel.WriteStringAndInt((int)ExcelColumns.AC, (int)ExcelColumns.AB, unique5WordsExact, 3);
                            }
                        }

                    }


                    uniqueWords = uniqueWords.Where(s1 => s1.number > 0).ToList();
                    uniqueWords.Sort((s1, s2) => s2.number.CompareTo(s1.number));
                    unique2WordsBroad = unique2WordsBroad.Where(s1 => s1.number > 0).ToList();
                    unique2WordsBroad.Sort((s1, s2) => s2.number.CompareTo(s1.number));
                    unique2WordsExact = unique2WordsExact.Where(s1 => s1.number > 0).ToList();
                    unique2WordsExact.Sort((s1, s2) => s2.number.CompareTo(s1.number));



                    generateRoots.Text = "Writing roots...";

                    writeExcel.WriteStringAndInt((int)ExcelColumns.E, (int)ExcelColumns.D, uniqueWords, 3);
                    writeExcel.WriteStringAndInt((int)ExcelColumns.H, (int)ExcelColumns.G, unique2WordsBroad, 3);
                    writeExcel.WriteStringAndInt((int)ExcelColumns.K, (int)ExcelColumns.J, unique2WordsExact, 3);

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