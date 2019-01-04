using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using KLAPT;
using LiteDB;

namespace KLAPT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string Sourcexlpath;
        public string Sourcewdpath;
        public string[] Splitted;
        public IEnumerable<WordList> Xldata;
        public IEnumerable<WordList> Falewordsindoc;



        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            var of = new OpenFileDialog {Filter = @"Excel Files|*.xls;*.xlsx;*.xlsm"};

            if (of.ShowDialog(this) == DialogResult.OK)
            {
                using (var db = new LiteDatabase(AppDomain.CurrentDomain.BaseDirectory + @"wordsData.db"))
                {
                    Sourcexlpath = of.InitialDirectory + of.FileName;
                    Xldata = GetDataTableFromExcel(Sourcexlpath).AsEnumerable().Select(
                        row => new WordList
                        {
                            FalseWord = row.Field<string>(0),
                            TrueWord = row.Field<string>(1)
                        }
                    ).ToList();

                    var col = db.GetCollection<WordList>("wordlist");
                    col.Insert(Xldata);

                    //dataGridView1.DataSource = col;

                    Xldata = col.FindAll().ToArray();
                    dataGridView1.DataSource = Xldata;
                    this.dataGridView1.Columns[0].Visible = false;

                    this.dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    this.dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                }

            }
        }

        public static string RemoveSpecialCharacters( string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if (char.IsNumber(c) || char.IsLetter(c) || c == '\'' || c == ' ')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }


        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            var of = new OpenFileDialog { Filter = @"txt Files|*.doc;*.docx;*.txt" };

            if (of.ShowDialog(this) == DialogResult.OK)
            {
                Sourcewdpath = of.InitialDirectory + of.FileName;
                var doc = new Document(Sourcewdpath);
                richTextBox1.Text = doc.GetText();

                var words = RemoveSpecialCharacters(richTextBox1.Text);
                Splitted = words.Split(' ');

                
                

            }

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
           
                var allxlsfalsewords = Xldata.Select(r => r.FalseWord).ToArray();
                string[] falewordsindoc = allxlsfalsewords.Intersect(Splitted).ToArray();



                Falewordsindoc = Xldata.Where(r => falewordsindoc.Contains(r.FalseWord));

                listBox1.Items.Clear();
                var falsewordLists = Falewordsindoc as WordList[] ?? Falewordsindoc.ToArray();
                listBox1.Items.AddRange(falsewordLists.Select(q =>
                    new
                    {
                        S = q.FalseWord + " ====> " + q.TrueWord

                    }).Select(a => a.S).ToArray());

                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Font.Bidi = true;
                builder.CurrentParagraph.ParagraphFormat.Bidi = true;
                builder.Font.LocaleIdBi = 1025;

                Aspose.Words.Font font = builder.Font;
                builder.Font.NameBi = "Arial";
                builder.Font.SizeBi = 16;

                builder.Write(richTextBox1.Text);
                
                var options = new FindReplaceOptions
                {
                    FindWholeWordsOnly = true,
                    ReplacingCallback = new ReplaceEvaluatorFindAndHighlight(),
                Direction = FindReplaceDirection.Backward
                };

                foreach (WordList falsewarod in falsewordLists)
                {
                    doc.Range.Replace(falsewarod.FalseWord, "", options);
                }
                string temppath = Path.GetDirectoryName(Sourcewdpath) +"\\" + Path.GetFileNameWithoutExtension(Sourcewdpath) + "-temp.rtf";
                builder.Document.Save(temppath , SaveFormat.Rtf);
            
                richTextBox1.LoadFile(temppath ) ;

            File.Delete(temppath);





        }

        private void Form1_Load(object sender, EventArgs e)
        {
            using (var db = new LiteDatabase(AppDomain.CurrentDomain.BaseDirectory + @"wordsData.db"))
            {
                // Get a collection (or create, if doesn't exist)
                var col = db.GetCollection<WordList>("wordlist");

                if (col.Count() > 0)
                {
                    Xldata = col.FindAll().ToArray();
                    dataGridView1.DataSource = Xldata.ToList();
                    this.dataGridView1.Columns[0].Visible = false;
                    this.dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    this.dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }

                

            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

            try
            {
                Document doc = new Document(Sourcewdpath);
                var options = new FindReplaceOptions
                {
                    FindWholeWordsOnly = true,
                    Direction = FindReplaceDirection.Forward
            };

                foreach (WordList falsewarod in Falewordsindoc)
                {
                    doc.Range.Replace(falsewarod.FalseWord, falsewarod.TrueWord, options);
                }

                doc.Save(Path.GetDirectoryName(Sourcewdpath) +"\\"+ Path.GetFileNameWithoutExtension(Sourcewdpath) + "-corrected" + Path.GetExtension(Sourcewdpath));

                MessageBox.Show(@"Done and corrected file saved");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
        private class ReplaceEvaluatorFindAndHighlight : IReplacingCallback
        {
            /// <summary>
            /// This method is called by the Aspose.Words find and replace engine for each match.
            /// This method highlights the match string, even if it spans multiple runs.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = e.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // In this case it is necessary to split the run.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);

                // This array is used to store all nodes of the match for further highlighting.
                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string.
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength = remainingLength - currentNode.GetText().Length;

                    // Select the next Run node. 
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    }
                    while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                // Split the last run that contains the match if there is any text left.
                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                // Now highlight all runs in the sequence.
                foreach (Run run in runs)
                    run.Font.HighlightColor = Color.Yellow;

                // Signal to the replace engine to do nothing because we have already done all what we wanted.
                return ReplaceAction.Skip;
            }
        }

        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }

    }
}
