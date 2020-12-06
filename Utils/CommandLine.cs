using System;
using System.Diagnostics;
using System.IO;

namespace WatcherConverter.Utils
{
    class CommandLine
    {
        public static int RunLibreOfficeConverter(string target, string outdir, string convertTo = "pdf")
        {
            Process p = new Process();
            string processName = "soffice.exe";
            string commandArgs = string.Format("--headless --convert-to {2} --outdir \"{1}\" \"{0}\"",
                target, outdir, convertTo);
            string libreOfficeDir = Environment.GetEnvironmentVariable("libreoffice");

            ProcessStartInfo s = new ProcessStartInfo(processName, commandArgs);

            s.WindowStyle = ProcessWindowStyle.Hidden;
            s.CreateNoWindow = true;
            s.UseShellExecute = true;
            p.StartInfo = s;

            s.WorkingDirectory = libreOfficeDir;
            //verify outdir exist
            DirectoryInfo directory = new DirectoryInfo(outdir);
            if (!directory.Exists)
            {
                directory.Create();
            }

            p.Start();
            p.WaitForExit();
            
            GC.Collect();
            return p.ExitCode;
        }
    }
}
