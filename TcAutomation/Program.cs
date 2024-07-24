using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NDesk.Options;
using System.IO;
using TCatSysManagerLib;

namespace TcAutomaticDeployment
{
    class Program
    {
        static void Main(string[] args)
        {
            //Placeholder for TwinCat source file path.
            string TcSourceFilePath = "";

            //Placeholder for AMS route ID.
            string amsNetId = "";

            //Shutdown help menu at startup.
            bool showHelp = false;

            //Options the user have, when executing the exe.
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
                Console.WriteLine("Usage: TcAutomatic Deployment [OPTIONS]");
                options.WriteOptionDescriptions(Console.Out);
                Environment.Exit(1);
            }

            //If file path is wrong or does not contain a TwinCat solution.
            if (!File.Exists(TcSourceFilePath))
            {
                Console.WriteLine("TwinCat solution does not exist!");
                Environment.Exit(1);
            }

            //If no AMS route adress has been provided, the AMS route will be set to local.
            if (string.IsNullOrEmpty(amsNetId))
            {
                Console.WriteLine("No AmsNetId provided, assuming local AmsNetId");
                amsNetId = "127.0.0.1.1.1";
            }

            //DTE version number
            Type t = System.Type.GetTypeFromProgID("TcXaeShell.DTE.17.0");
            EnvDTE.DTE dte = (EnvDTE.DTE)System.Activator.CreateInstance(t);

            //Suppress the TwinCat developer window, so it wont take resources. 
            //If you want to see the TwinCat developer window, change "dte.SuppressUI = false", and "dte.MainWindow.Visible = true"
            dte.SuppressUI = true;
            dte.MainWindow.Visible = false;

            //Passes the solution given by the file path
            EnvDTE.Solution sol = dte.Solution;
            sol.Open(TcSourceFilePath);

            EnvDTE.Project pro = sol.Projects.Item(1);


            ITcSysManager15 sysMan = (ITcSysManager15) pro.Object;

            //Sets the AMS route adress
            sysMan.SetTargetNetId(amsNetId);

            //Activates the configuration of the TwinCat solution on the AMS route adress.
            sysMan.ActivateConfiguration();

            //Reboots the TwinCat PLC
            sysMan.StartRestartTwinCAT();

            Console.WriteLine("Done!");
                       

        }
    }
}
