using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;

namespace MxDesktopAppHelpers
{
    public class AcadConsoleProcess
    {
        private static string AcadConsoleExePath
        {
            get
            {
                return @"C:\Program Files\Autodesk\AutoCAD 2019\accoreconsole.exe";
            }
        }

        private static string CreateScriptFile(params string[] commands)
        {
            // build a text with the command list
            // one command per line, no spaces (Trim)
            StringBuilder listOfCommands = new StringBuilder();
            foreach (string command in commands)
            {
                if (command.Contains(" "))
                {
                    listOfCommands.AppendLine(string.Format("\"{0}\"", command));
                }
                else
                {
                    listOfCommands.AppendLine(command.Trim());
                }
            }

            // ensure AutoCAD Quit at the end
            listOfCommands.AppendLine("_.QUIT");

            // unique script file name
            string scrFileName =
              Path.GetTempPath() +
              Guid.NewGuid().ToString() +
              ".scr";
            File.WriteAllText(scrFileName, listOfCommands.ToString());

            return scrFileName;
        }

        public static string RunCommands(string fileName, params string[] commands)
        {
            // create a script file with the commands
            string scrFileName = CreateScriptFile(commands);
            string output = string.Empty;
            try
            {
                output = RunCommands(fileName, scrFileName);
            }
            finally
            {
                // erase the temp script file
                File.Delete(scrFileName);
            }
            return output;
        }

        public static string RunCommands(string fileName, string scriptFile)
        {
            bool redirectOutput = false;
            // no window and redirect the output
            Process process = new Process();
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = redirectOutput;
            process.StartInfo.CreateNoWindow = false;

            // parameters to execute the script file
            process.StartInfo.FileName = AcadConsoleExePath;

            // build the parameters
            StringBuilder param = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(scriptFile))
                param.AppendFormat(" /s \"{0}\"", scriptFile);
            if (!string.IsNullOrWhiteSpace(fileName))
                param.AppendFormat(" /i \"{0}\"", fileName);

            process.StartInfo.Arguments = param.ToString();

            string output = string.Empty;
            try
            {
                using (process)
                {
                    // run it!
                    process.Start();

                    // read the output to return
                    // this will stop this execute until AutoCAD exits
                    if (redirectOutput)
                    {
                        StreamReader outputStream = process.StandardOutput;
                        output = outputStream.ReadToEnd();
                        outputStream.Close();
                    }                    
                }
            }
            catch (Exception ex)
            {
                output = ex.Message;
            }

            return output.Replace("\0", string.Empty);
        }
    }
}
