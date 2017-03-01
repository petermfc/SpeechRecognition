using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace SpeechRecognition
{
    static class Program
    {
        public static Assembly ribbon = null;
        public static Assembly nAudio = null;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            LoadAssemblies();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        private static void LoadAssemblies()
        {
            string resourceRibbon = Application.ProductName + ".System.Windows.Forms.Ribbon35.dll";
            using (Stream stm = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceRibbon))
            {
                byte[] ba = new byte[(int)stm.Length];
                stm.Read(ba, 0, (int)stm.Length);
                ribbon = Assembly.Load(ba);
            }
            string resourceNAudio = Application.ProductName + ".NAudio.dll";
            using (Stream stm = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceNAudio))
            {
                byte[] ba = new byte[(int)stm.Length];
                stm.Read(ba, 0, (int)stm.Length);
                nAudio = Assembly.Load(ba);
            }

            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            if (args.Name.StartsWith("NAudio"))
            {
                return nAudio;
            }
            else if (args.Name.StartsWith("System.Windows.Forms.Ribbon35"))
            {
                return ribbon;
            }
            return null;
        }
    }
}
