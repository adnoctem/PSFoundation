using System;
using System.Diagnostics;
using System.Text;
using System.Threading;

public static class NativeProcess
{
    public static int Main(string[] args)
    {
        switch (args[0])
        {
            case "arguments":
                for (int i = 1; i < args.Length; i++)
                    Console.WriteLine("ARG:" + Convert.ToBase64String(Encoding.UTF8.GetBytes(args[i])));
                return 0;
            case "streams":
                for (int i = 0; i < 512; i++)
                {
                    Console.Out.Write(new string('O', 1024));
                    Console.Error.Write(new string('E', 1024));
                }
                return 17;
            case "wait":
                Console.WriteLine(Process.GetCurrentProcess().Id);
                Console.Out.Flush();
                Thread.Sleep(60000);
                return 0;
            case "tree":
                using (Process child = Process.Start(new ProcessStartInfo {
                    FileName = Process.GetCurrentProcess().MainModule.FileName,
                    Arguments = "sleep", UseShellExecute = false, CreateNoWindow = true
                }))
                {
                    Console.WriteLine(Process.GetCurrentProcess().Id);
                    Console.WriteLine(child.Id);
                    Console.Out.Flush();
                    Thread.Sleep(60000);
                }
                return 0;
            case "sleep":
                Thread.Sleep(60000);
                return 0;
            default:
                return 2;
        }
    }
}
