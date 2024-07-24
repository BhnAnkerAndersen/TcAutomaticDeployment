using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NDesk.Options;
using System.IO;
using TCatSysManagerLib;

namespace TcAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            string TcSourceFilePath = "";
            string amsNetId = "";
            bool showHelp = false;


            OptionSet options = new OptionSet()
           .Add("h|?|help", delegate (string v) { showHelp = v != null; })
           .Add("f|tcFilePath=", delegate (string v) { TcSourceFilePath = v; })
           .Add("a|amsNetId=", delegate (string v) { amsNetId = v; });
       

            try
            {
                options.Parse(args);
            }
            catch (OptionException e)
            {
                Console.WriteLine(e.Message);
                Environment.Exit(1);
            }

            if (showHelp)
            {
                Console.WriteLine("Usage: TcAutomation [OPTIONS]");
                options.WriteOptionDescriptions(Console.Out);
                Environment.Exit(1);
            }

            if (!File.Exists(TcSourceFilePath))
            {
                Console.WriteLine("TwinCat solution does not exist!");
                Environment.Exit(1);
            }

            if (string.IsNullOrEmpty(amsNetId))
            {
                Console.WriteLine("No AmsNetId provided, assuming local AmsNetId");
                amsNetId = "127.0.0.1.1.1";
            }

            Type t = System.Type.GetTypeFromProgID("TcXaeShell.DTE.15.0");
            EnvDTE.DTE dte = (EnvDTE.DTE)System.Activator.CreateInstance(t);
            dte.SuppressUI = false;
            dte.MainWindow.Visible = true;

            EnvDTE.Solution sol = dte.Solution;
            sol.Open(TcSourceFilePath);

            EnvDTE.Project pro = sol.Projects.Item(1);


            ITcSysManager15 sysMan = (ITcSysManager15) pro.Object;


            sysMan.SetTargetNetId(amsNetId);
            sysMan.ActivateConfiguration();
            sysMan.StartRestartTwinCAT();

            Console.WriteLine("Done!");
                       

        }
    }
}
