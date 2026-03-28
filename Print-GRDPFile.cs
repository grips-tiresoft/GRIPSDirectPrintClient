using System;
using System.Diagnostics;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 1)
            return;

        string exeDir = AppDomain.CurrentDomain.BaseDirectory;
        string psScript = Path.Combine(exeDir, "Print-GRDPFile.ps1");
        string inputFile = args[0];

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = string.Format("-STA -WindowStyle Hidden -NoProfile -ExecutionPolicy RemoteSigned -File \"{0}\" \"{1}\"", psScript, inputFile),
            WindowStyle = ProcessWindowStyle.Hidden,
            CreateNoWindow = true,
            UseShellExecute = false
        };

        Process.Start(psi);
    }
}