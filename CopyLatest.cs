using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace copyonlylatest
{
    class Program
    {
        static void Main(string[] args)
        {

            string src = System.Configuration.ConfigurationManager.AppSettings["src"];
            string srcHF = System.Configuration.ConfigurationManager.AppSettings["srcHF"];
            string apuPath = System.Configuration.ConfigurationManager.AppSettings["APUpath"];
            string symbolPath = System.Configuration.ConfigurationManager.AppSettings["symbolPath"];
            string dst = System.Configuration.ConfigurationManager.AppSettings["dst"];
            string dstsym = System.Configuration.ConfigurationManager.AppSettings["dstsymbol"];
            string indipath = System.Configuration.ConfigurationManager.AppSettings["ipath"];
            string attempts = System.Configuration.ConfigurationManager.AppSettings["attempts"];
            string copyhf = System.Configuration.ConfigurationManager.AppSettings["CopyHF"];
            string copyapu = System.Configuration.ConfigurationManager.AppSettings["CopyAPU"];
            int lastcount = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["lastcount"]);
            int chkdays = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["Hourstocheck"]);
            var directory = new DirectoryInfo(src);
            var directory1 = new DirectoryInfo(srcHF);
            var directory2 = new DirectoryInfo(apuPath);
            DateTime dtnow = DateTime.Now;
            DateTime dt3dbfr = dtnow.Subtract(TimeSpan.FromHours(chkdays));
            #region clientbuildand UpdateMamnger
            List<string> todayfiles = (from f in directory.GetFiles()
                                       orderby f.LastWriteTime descending
                                       where f.CreationTime > dt3dbfr
                                       select f.Name).ToList();

            if (todayfiles.Count == 0)
            {
                LogMessage("No Latest Files were Avialble whne checked this time: between:   " + dtnow.ToString() + "  and " + dt3dbfr.ToString());
            }
            else
            {
                LogMessage("Today files count whne checked this time: between :   " + dtnow.ToString() + "  and " + dt3dbfr.ToString() + "  " + todayfiles.Count.ToString());
                foreach (var iii in todayfiles)
                {
                    LogMessage("File anme " + iii);
                }
            }
            LogMessage("Source Location for Builds ,CDROM Desktop client,Admin Cleint,Updatemanagerclient is  =: " + src);
            foreach (var indfile in todayfiles)
            {
               // if (!indfile.Contains("LowisSetup") && !indfile.Contains("LowisAdminSetup"))
               // {
                    Console.WriteLine("File Name : = " + indfile);
                    LogMessage("File Name : = " + indfile);
                    if (indfile.Length > 0)
                    {
                        dorobocopy(src, dst, indfile);
                    }
              //  }

            }

            #endregion
            #region HotFixFolders
            if (copyhf.ToLower() == "y")
            {
                List<string> todayhffols = (from f1 in directory1.GetDirectories()
                                            orderby f1.LastWriteTime descending
                                            where f1.CreationTime > dt3dbfr
                                            select f1.Name).ToList();
                LogMessage("Hotfixes Location: " + srcHF);
                if (todayhffols.Count == 0)
                {
                    LogMessage("No Latest HF folders were Avialble whenchecked this time ");
                }
                foreach (var indfile1 in todayhffols)
                {

                    Console.WriteLine("Hotfixes folders name : = " + indfile1);
                    LogMessage("Hotfixes folders name : = " + indfile1);
                    var directoryinddest = new DirectoryInfo(Path.Combine(dst, indfile1));
                    if (!directoryinddest.Exists)
                    {
                        directoryinddest.Create();
                    }
                    if (indfile1 != "APUs")
                    {
                        dorobocopyFolders(Path.Combine(srcHF, indfile1), directoryinddest.ToString());
                    }

                }
            }
            #endregion
            #region APUS
            //if (todayhffols.Count == 0)
            //{
            //    LogMessage("No LatestAPU folders were Avialble whenchecked this time ");
            //}
            //Will copy Latest APU folders Today..only Needs to run everyday
            LogMessage("APU location: " + apuPath);
            if (copyapu.ToLower() == "y")
            {
                List<string> todayapus = (from f2 in directory2.GetDirectories()
                                          orderby f2.LastWriteTime descending
                                          where f2.CreationTime > dt3dbfr
                                          select f2.Name).ToList();

                foreach (var indfile2 in todayapus)
                {

                    Console.WriteLine("APU folders name : = " + indfile2);
                    LogMessage("APU folders name : = " + indfile2);
                    dorobocopyFolders(Path.Combine(apuPath, indfile2), dst);

                }
            }

            #endregion
            #region Symbolfiles
          //  Console.WriteLine("Loking for symbol files");
          //  dorobocopyFolders(symbolPath,dstsym);
            #endregion
            //   Console.ReadLine();

        }
        #region RequiredFunctions
        private static void dorobocopy(string src, string dst, string fln)
        {
            LogMessage("File Name : = " + fln);
            if (src != "" && dst != "" && fln != "")
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = false;
                startInfo.FileName = "robocopy.exe";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Arguments = " \"" + src + "\"" + " \"" + dst + "\"" + " \"" + fln + "\"" + " /z";
                try
                {
                    using (Process exeProcess = Process.Start(startInfo))
                    {
                        exeProcess.WaitForExit();
                    }
                }
                catch
                {
                    // Log error.
                }
            }
        }
        private static void dorobocopyFolders(string src, string dst)
        {

            if (src != "" && dst != "")
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = false;
                startInfo.FileName = "robocopy.exe";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Arguments = "  \"" + src + "\"" + " \"" + dst + "\"" + " *.* /e /z ";
                try
                {
                    using (Process exeProcess = Process.Start(startInfo))
                    {
                        exeProcess.WaitForExit();
                    }
                }
                catch
                {
                    // Log error.
                }
            }
        }
        private static void LogMessage(string txt)
        {
            System.IO.File.AppendAllText(@"E:\logtsk.txt", "[" + System.DateTime.Now.ToString() + "] : " + txt + Environment.NewLine);
        }
        #endregion
    }
}
