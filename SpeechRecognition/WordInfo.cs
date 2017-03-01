using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeechRecognition
{
    public class WordInfo
    {
        private static long counter = 0L;

        private static long GetNextId()
        {
            return ++counter;
        }

        private WordInfo() { }

        private WordInfo(string text, string confidence, string lexicalForm, bool uncertain)
        {
            this.Id = GetNextId();
            this.Text = text;
            this.Confidence = confidence;
            this.LexicalForm = lexicalForm;
            this.Uncertain = uncertain;
        }

        public static WordInfo Create(string text, string confidence)
        {
            return new WordInfo(text, confidence, "", false);
        }

        public static WordInfo Create(string text, string confidence, string lexicalForm)
        {
            return new WordInfo(text, confidence, lexicalForm, !String.IsNullOrEmpty(lexicalForm));
        }

        public void SetStartPosition(int pos)
        {
            this.StartPosition = pos;
        }

        public void SetStopPosition(int pos)
        {
            this.StopPosition = pos;
        }

        public long Id { get; private set; }
        public string Text { get; set; }
        public string Confidence { get; private set; }
        public string LexicalForm { get; private set; }
        public bool Uncertain { get; private set;}
        public int StartPosition { get; private set; }
        public int StopPosition { get; private set; }

    }
}
