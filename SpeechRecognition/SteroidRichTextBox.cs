using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpeechRecognition
{
    class SteroidRichTextBox: RichTextBox
    {
        public static readonly string NEWLINE = "\n";
        private ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
        private ToolStripItem miConfidence = new ToolStripMenuItem();
        private List<KeyValuePair<TimeSpans, List<WordInfo>>> wordLines = new List<KeyValuePair<TimeSpans, List<WordInfo>>>();
        private StringBuilder sb = new StringBuilder();

        public bool AutoScroll { get; internal set; }

        public SteroidRichTextBox()
            :base()
        {
            this.ContextMenuStrip = contextMenuStrip;
            this.MouseDown += SteroidRichTextBox_MouseDown;
        }

        private void SteroidRichTextBox_MouseDown(object sender, MouseEventArgs e)
        {
            MouseEventArgs me = e as MouseEventArgs;
            if(me != null)
            {
                if(me.Button == MouseButtons.Right)
                {
                    contextMenuStrip.Items.Clear();
                    WordInfo selectedWord = GetSelectedWord(e.Location);
                    if (selectedWord != null)
                    {
                        if (!selectedWord.Uncertain)
                        {
                            miConfidence.Text = selectedWord.Text + " : " + selectedWord.Confidence;
                            miConfidence.Enabled = false;
                            contextMenuStrip.Items.Add(miConfidence);
                        }
                        else
                        {
                            miConfidence.Text = "Confidence: " + selectedWord.Confidence + "%";
                            miConfidence.Enabled = false;
                            contextMenuStrip.Items.Add(miConfidence);
                            contextMenuStrip.Items.Add("-");
                            ToolStripItem miUncertain = new ToolStripMenuItem();
                            miUncertain.Text = selectedWord.LexicalForm;
                            miUncertain.Tag = selectedWord;
                            contextMenuStrip.Items.Add(miUncertain);
                            miUncertain.Click += MiUncertain_Click;
                        }
                    }
                }
            }
        }

        private void MiUncertain_Click(object sender, EventArgs e)
        {
            ToolStripItem tsi = sender as ToolStripMenuItem;
            if(tsi != null)
            {
                WordInfo wi = tsi.Tag as WordInfo;
                if(wi != null)
                {
                    int oldLen = wi.Text.Length;
                    int newLen = tsi.Text.Length;
                    string oldWord = wi.Text;
                    string newWord = tsi.Text.PadRight(oldLen - newLen);
                    WordInfo wiToChange = wordLines.SelectMany(s => s.Value).Where(w => w.Id == wi.Id).SingleOrDefault();
                    wiToChange.Text = newWord;
                    RefreshText();
                }
            }
        }

        private WordInfo GetSelectedWord(Point location)
        {
            int pos = this.GetCharIndexFromPosition(new Point(location.X, location.Y));
            int lineIndex = this.GetLineFromCharIndex(pos);
            var line = this.wordLines[lineIndex];
            WordInfo wi = line.Value.Where(w => w.StartPosition <= pos &&  pos <= w.StopPosition).FirstOrDefault();
            return wi;
        }

        public void AddLine(TimeSpan timeSpan, IEnumerable<WordInfo> wordInfos)
        {
            TimeSpans timeSpans = TimeSpans.Create(timeSpan, timeSpan);
            if(wordLines.Any())
            {
                timeSpans.StartTime = wordLines.Last().Key.StopTime;
            }
            else
            {
                timeSpans = TimeSpans.Create(new TimeSpan(0), new TimeSpan(0));
            }
            var line = new KeyValuePair<TimeSpans, List<WordInfo>>(timeSpans, new List<WordInfo>());
            line.Value.AddRange(wordInfos);
            wordLines.Add(line);
            string time = String.Format("[{0:D2}:{1:D2}.{2:D2}]", timeSpans.StartTime.Minutes, timeSpans.StartTime.Seconds, timeSpans.StartTime.Milliseconds);
            sb.Append(time).Append(": ");
            AppendWordInfoCollection(sb, wordInfos, wordLines.Count - 1);
            sb.Append(NEWLINE);
        }

        private static void AppendWordInfoCollection(StringBuilder sb, IEnumerable<WordInfo> wordInfos, int offset)
        {
            foreach(WordInfo wi in wordInfos)
            {
                wi.SetStartPosition(sb.Length - offset);
                sb.Append(wi.Text);
                wi.SetStopPosition(sb.Length - offset);
                sb.Append(" ");
            }
        }

        public void RefreshText()
        {
            sb.Clear();
            for(int i = 0; i < wordLines.Count; i++)
            {
                var line = wordLines[i];
                string time = String.Format("[{0:D2}:{1:D2}.{2:D2}]", line.Key.StartTime.Minutes, line.Key.StartTime.Seconds, line.Key.StartTime.Milliseconds);
                sb.Append(time).Append(": ");
                AppendWordInfoCollection(sb, line.Value, i);
                sb.Append(NEWLINE);
            }

            this.Text = sb.ToString();
            this.SelectionStart = sb.Length - 1;
            if(AutoScroll)
            {
                this.ScrollToCaret();
            }
        }

        public new void Clear()
        {
            base.Clear();
            wordLines.Clear();
            sb.Clear();
        }
    }
}
