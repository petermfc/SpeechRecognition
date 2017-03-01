using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Speech.Recognition;
using System.Text;
using System.Windows.Forms;
using System.Linq;

namespace SpeechRecognition
{
    public partial class MainForm : Form
    {
        private SpeechRecognitionEngine recognitionEngine;
        private float confidence = 0.7f;
        private WaveStream audioStream;
        private RecognizerInfo recognizerInfo = null;
        public MainForm()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(MainForm_DragEnter);
            this.DragDrop += new DragEventHandler(MainForm_DragDrop);
            this.txtOutput.EnableAutoDragDrop = false;
            Init();
        }

        private void Init()
        {
            SetAvailableLanguages();

            SetDoubleBuffering(txtOutput);
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabelPosition.Visible = false;
            toolStripStatusLabelStatus.Text = "Select file";
            btnRestart.Enabled = false;
            btnStart.Enabled = false;
            btnStop.Enabled = false;
            txtOutput.Clear();

        }

        private void SetAvailableLanguages()
        {
            List<ToolStripItem> items = new List<ToolStripItem>();
            foreach (RecognizerInfo ri in SpeechRecognitionEngine.InstalledRecognizers())
            {
                ToolStripMenuItem ti = new ToolStripMenuItem();
                ti.Text = ri.Culture.EnglishName;
                ti.CheckOnClick = true;
                if (ri.Culture.TwoLetterISOLanguageName.Equals("en"))
                {
                    ti.Checked = true;
                    recognizerInfo = ri;
                }
                else
                {
                    ti.Checked = false;
                }
                ti.CheckedChanged += languageToolStripMenuItem_CheckedChanged;
                items.Add(ti);
            }
            items.Add(toolStripSeparator);
            items.Add(installMoreLanguagesToolStripMenuItem);
            languageToolStripMenuItem.DropDownItems.Clear();
            languageToolStripMenuItem.DropDownItems.AddRange(items.ToArray());
        }

        private void SetUpWithFile(string filePath)
        {
            RecognizerInfo info = null;
            foreach (RecognizerInfo ri in SpeechRecognitionEngine.InstalledRecognizers())
            {
                if (ri.Culture.TwoLetterISOLanguageName.Equals("en"))
                {
                    info = ri;
                    break;
                }
            }
            if (info == null) return;

            /*Set up audio*/
            audioStream = new WaveFileReader(filePath);
            TimeSpan audioStreamTotalTime = audioStream.TotalTime;

            // Create the selected recognizer.
            recognitionEngine = new SpeechRecognitionEngine(info);
            recognitionEngine.LoadGrammar(new DictationGrammar());
            recognitionEngine.SetInputToWaveFile(filePath);
            StringBuilder sb = new StringBuilder();
            recognitionEngine.SpeechRecognized += (s, args) =>
            {
                TimeSpan currentTime = new TimeSpan(recognitionEngine.RecognizerAudioPosition.Ticks);
                string time = String.Format("[{0:D2}:{1:D2}.{2:D2}]", currentTime.Minutes, currentTime.Seconds, currentTime.Milliseconds);

                /*Update progress label*/
                String positionString = String.Format("{0} / {1}", currentTime, audioStreamTotalTime);
                toolStripStatusLabelPosition.Text = positionString;
                double progress = (double)currentTime.Ticks / (double)audioStreamTotalTime.Ticks * 100.0;
                toolStripProgressBar1.Value = (int)progress;

                List<WordInfo> wordInfos = new List<WordInfo>();
                foreach (RecognizedWordUnit word in args.Result.Words)
                {
                    WordInfo wi;
                    string confidenceStr = String.Format("{0:0}", word.Confidence * 100);
                    if (word.Confidence >= confidence)
                    {
                        wi = WordInfo.Create(word.Text, confidenceStr, word.LexicalForm);
                    }
                    else
                    {
                        string text = String.Format("[SKIPPED]", confidenceStr);
                        wi = WordInfo.Create(text, confidenceStr, word.LexicalForm);
                    }
                    wordInfos.Add(wi);
                }
                txtOutput.AddLine(currentTime, wordInfos);
                txtOutput.RefreshText();
            };
            recognitionEngine.RecognizeCompleted += RecognitionEngine_RecognizeCompleted;
            btnStart.Enabled = true;
            toolStripStatusLabelStatus.Text = "File loaded: " + filePath;
            lblConfValue.Text = String.Format("{0:P2}", confidence);
        }

        private void RecognitionEngine_RecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            btnRestart.Enabled = true;
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabelStatus.Text = "All recognized";
        }

        private void SetDoubleBuffering(Control c)
        {
            if (System.Windows.Forms.SystemInformation.TerminalServerSession)
                return;

            System.Reflection.PropertyInfo aProp =
                  typeof(System.Windows.Forms.Control).GetProperty(
                        "DoubleBuffered",
                        System.Reflection.BindingFlags.NonPublic |
                        System.Reflection.BindingFlags.Instance);

            aProp.SetValue(c, true, null);
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            toolStripStatusLabelPosition.Visible = true;
            toolStripProgressBar1.Visible = true;
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = 100;
            recognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
            btnRestart.Enabled = false;
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            toolStripStatusLabelStatus.Text = "Recognition in progress!";
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabelPosition.Visible = true;
            recognitionEngine.RecognizeAsyncStop();
            btnRestart.Enabled = true;
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            toolStripStatusLabelStatus.Text = "Recognition Stopped";
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            confidence = (float)trackBar1.Value / 100;
            lblConfValue.Text = String.Format("{0:P2}", confidence);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (audioStream != null)
            {
                audioStream.Close();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            if(sender is SteroidRichTextBox)
            {
                HandleDrag(e);
                return;
            }
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            HandleDrag(e);
        }

        private void HandleDrag(DragEventArgs e)
        {
            if (e.Data != null)
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1)
                {
                    if (System.IO.Path.GetExtension(files[0]).Equals(".wav", StringComparison.InvariantCultureIgnoreCase))
                    {
                        SetUpWithFile(files[0]);
                    }
                }
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SetUpWithFile(openFileDialog1.FileName);
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.AddExtension = true;
            sfd.DefaultExt = "txt";
            sfd.SupportMultiDottedExtensions = false;
            sfd.OverwritePrompt = true;
            sfd.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            sfd.FileName = "RecognizedText.txt";

            if(sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    StringBuilder stringBuilder = new StringBuilder(txtOutput.Text);
                    stringBuilder.Replace("]: ", "]");
                    using (StreamWriter sw = new StreamWriter(sfd.FileName))
                    {
                        stringBuilder.ToString().Split(SteroidRichTextBox.NEWLINE.ToCharArray()).ToList()
                            .ForEach(p =>
                                sw.WriteLine(p.TrimEnd())
                            );
                    }
                    toolStripStatusLabelStatus.Text = "File saved: " + sfd.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Error during file saving.\n" + ex.StackTrace);
                    toolStripStatusLabelStatus.Text = "Error during file saving.";
                }
            }
        }

        private void languageToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            var currentItem = sender as ToolStripMenuItem;
            if (currentItem != null)
            {
                ToolStripItemCollection itemCollection = ((ToolStripMenuItem)currentItem.OwnerItem).DropDownItems;
                if (itemCollection.Count > 2)
                {
                    foreach (ToolStripItem item in itemCollection)
                    {
                        ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                        if (menuItem != null)
                        {
                            menuItem.CheckedChanged -= languageToolStripMenuItem_CheckedChanged;
                            menuItem.Checked = false;
                            menuItem.CheckedChanged += languageToolStripMenuItem_CheckedChanged;
                        }
                    }
                    currentItem.CheckedChanged -= languageToolStripMenuItem_CheckedChanged;
                    currentItem.Checked = true;
                    currentItem.CheckedChanged += languageToolStripMenuItem_CheckedChanged;
                }
            }
        }

        private void installMoreLanguagesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Speech Recognition is available only in following language packs: English, French, Spanish, German, Japanese, Simplified Chinese, and Traditional Chinese.", "SpeechRecognition", MessageBoxButtons.OK, MessageBoxIcon.Information);
            var cplPath = System.IO.Path.Combine(Environment.SystemDirectory, "control.exe");
            System.Diagnostics.Process.Start(cplPath, "/name Microsoft.Language");
        }

        private void btnRestart_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Really?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(res == DialogResult.Yes)
            {
                Init();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Super program!\nVersion 0.2\n©2016, 2017 Piotr Darosz", Application.ProductName);
        }

        private void chAutoScroll_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;
            if (checkBox != null) {
                txtOutput.AutoScroll = checkBox.Checked;
            }
        }
    }
}