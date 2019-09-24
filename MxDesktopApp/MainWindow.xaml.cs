using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace MxDesktopApp
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<string> DWGFilenames = new List<string>();

        readonly Process _cmd;

        private static StringBuilder sortOutput = null;


        public MainWindow()
        {
            InitializeComponent();
        }

        

        #region Buttons
        private void Button_LoadFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = @"C:\Temp\Metro";

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string filename in openFileDialog.FileNames)
                    //DWGFiles.Items.Add(System.IO.Path.GetFileName(filename));
                    //DWGFilenames.Add(System.IO.Path.GetFileName(filename));
                    DWGFilenames.Add(filename);
            }

            DWGFiles.ItemsSource = DWGFilenames;
        }

        private void ButtonClear_Click(object sender, RoutedEventArgs e)
        {
            DWGFiles.ClearValue(ItemsControl.ItemsSourceProperty);
            DWGFilenames.Clear();
            outputTextBox.Text = "Selection cleared";
            output.Text = "Output cleared";
            
        }

        private void ButtonDeleteBAK_Click(object sender, RoutedEventArgs e)
        {
            string file = DWGFilenames[0];
            string directory = Path.GetDirectoryName(file);

            List<FileInfo> files = GetFiles(directory, ".bak", ".pcp");

            int deleted = 0;
            foreach (FileInfo fi in files)
            {
                try
                {
                    fi.Attributes = FileAttributes.Normal;
                    File.Delete(fi.FullName);
                    deleted++;
                }
                catch { }
            }
            outputTextBox.Text = $"{deleted} files deleted";
        }

        private void ButtonRunNotOverlapping_Click(object sender, RoutedEventArgs e)
        {
            
            string script = Path.Combine(Environment.CurrentDirectory, "runNotOverlapping.scr");

            if (DWGFilenames.Count > 0)
            {
                output.Text = "Start processing Not Overlapping Sheets\n";

                //foreach (string filename in DWGFilenames)
                //{
                // //   prgBar.Value += 100 / DWGFilenames.Count;

                //    outputTextBox.Text += $"{filename}...{RunCommands(filename, script, outputTextBox, _cmd)}\n";

                //}

                Dictionary<string, string> list = GetEntityBreakUp(DWGFilenames,script,outputTextBox);

                //prgBar.Value = 100;
            }
            else
            {
                MessageBox.Show("No DWG selected");
            }
        }

        private void ButtonRunOverlapping_Click(object sender, RoutedEventArgs e)
        {
            
            string script = Path.Combine(Environment.CurrentDirectory, "runOverlapping.scr");

            if (DWGFilenames.Count > 0)
            {
                output.Text = "Start processing Overlapping Sheets\n";

                foreach (string filename in DWGFilenames)
                {
                    

                    outputTextBox.Text += $"{filename}...{RunCommands(filename, script, outputTextBox, _cmd)}\n";

                }

                
            }
            else
            {
                MessageBox.Show("No DWG selected");
            }
        }
        #endregion

        #region Utils
        private FileInfo[] GetDirectoryContent(string Folder, string FileType)
        {
            DirectoryInfo dinfo = new DirectoryInfo(Folder);
            return dinfo.GetFiles(FileType);

        }

        public List<FileInfo> GetFiles(string path, params string[] extensions)
        {
            List<FileInfo> list = new List<FileInfo>();
            foreach (string ext in extensions)
                list.AddRange(new DirectoryInfo(path).GetFiles("*" + ext).Where(p =>
                      p.Extension.Equals(ext, StringComparison.CurrentCultureIgnoreCase))
                      .ToArray());
            return list;
        }

        // Launches multiple instance of AccoreConsole
        public Dictionary<String, String> GetEntityBreakUp(List<string> fileNames, string scriptFile, TextBox tboxOutput)
        {
            object lockObject = new object();
            Dictionary<String, String> entBreakup = new Dictionary<String, String>();

            Parallel.ForEach(
                // The values to be aggregated
                fileNames,

                // The local initial partial result
                () => new Dictionary<String, String>(),

                // The loop body
                (x, loopState, partialResult) =>
                {
                    // Lauch AccoreConsole and find the entity breakup
                    //FileInfo fi = x as FileInfo;

                    String consoleOutput = String.Empty;
                    String entityBreakup = String.Empty;
                    using (Process coreprocess = new Process())
                    {
                        coreprocess.StartInfo.UseShellExecute = false;
                        coreprocess.StartInfo.CreateNoWindow = true;
                        coreprocess.StartInfo.RedirectStandardOutput = true;
                        coreprocess.StartInfo.FileName = @"C:\Program Files\Autodesk\AutoCAD 2019\accoreconsole.exe";

                        //coreprocess.StartInfo.Arguments = string.Format("/i \"{0}\" /s \"{1}\" /l en-US", fi.FullName, @"C:\Temp\RunCustomNETCmd.scr");

                        // parameters to execute the script file
                        coreprocess.StartInfo.Arguments = $"/i {x} /s \"{scriptFile}\" /l en-US";

                        coreprocess.Start();

                        // Max wait for 5 seconds 
                        coreprocess.WaitForExit(5000);



                        //StreamReader outputStream
                        //        = coreprocess.StandardOutput;
                        //consoleOutput = outputStream.ReadToEnd();

                        //String cleaned = consoleOutput.Replace("\0", string.Empty);

                        coreprocess.OutputDataReceived += new DataReceivedEventHandler(_cmd_OutputDataReceived);
                        

                        //async
                        coreprocess.BeginOutputReadLine();



                        //UpdateConsole(cleaned);

                        //outputStream.Close();
                        
                        

                    }

                    Dictionary<String, String> partialDict = partialResult as Dictionary<String, String>;

                    //partialDict.Add(x.FullName, entityBreakup);
                    partialDict.Add(x, entityBreakup);
                    return partialDict;
                },

                // The final step of each local context
                (partialEntBreakup) =>
                {
                    // Enforce serial access to single, shared result
                    lock (lockObject)
                    {
                        Dictionary<String, String> partialDict = partialEntBreakup as Dictionary<String, String>;

                        foreach (KeyValuePair<String, String> kvp in partialDict)
                        {
                            entBreakup.Add(kvp.Key, kvp.Value);
                            
                        }
                    }
                });

            
            return entBreakup;
        }

        public string RunCommands(string fileName, string scriptFile, TextBox txtResult, Process _cmd)
        {
            
            ProcessStartInfo startInfo = new ProcessStartInfo();
            // no window and redirect the output

            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.CreateNoWindow = true;

            sortOutput = new StringBuilder();

            //startInfo.StandardOutputEncoding = UnicodeEncoding.Default;

            // accoreconsole file path
            startInfo.FileName = @"C:\Program Files\Autodesk\AutoCAD 2019\accoreconsole.exe";

            // parameters to execute the script file
            startInfo.Arguments = $"/i {fileName} /s \"{scriptFile}\" /l en-US";

            // log result
            //string path = @"C:\Temp\Metro\Test\result.txt";

            string output = string.Empty;
            try
            {
                using (_cmd = new Process())
                {
                    _cmd.StartInfo = startInfo;

                    // run it
                    bool processStarted = _cmd.Start();

                    txtResult.Text = "Started...";

                    if (processStarted)
                    {

                        _cmd.OutputDataReceived += new DataReceivedEventHandler(_cmd_OutputDataReceived);
                        _cmd.ErrorDataReceived += new DataReceivedEventHandler(_cmd_ErrorDataReceived);
                        _cmd.Exited += new EventHandler(_cmd_Exited);

                        //async
                        _cmd.BeginOutputReadLine();
                        _cmd.BeginErrorReadLine();

                        //Get the output stream
                        //outputReader = process.StandardOutput;
                        //errorReader = process.StandardError;
                        //process.WaitForExit();

                        //Display the result
                        //string displayText = "Output" + Environment.NewLine + "==============" + Environment.NewLine;
                        //displayText += outputReader.ReadToEnd();
                        //displayText += Environment.NewLine + "Error" + Environment.NewLine + "==============" +
                        //               Environment.NewLine;
                        //displayText += errorReader.ReadToEnd();
                        //txtResult.Text = displayText;

                    }

                    if (_cmd.HasExited)
                    {
                        outputTextBox.Text = "done";
                    }

                    // read the output to return
                    // this will stop this execute until AutoCAD exits
                    //error cannot mix async and sync processes
                    //StreamReader outputStream = _cmd.StandardOutput;
                    //output = outputStream.ReadToEnd();
                    //outputStream.Close();

                    //File.AppendAllText(path, output.ToString());

                    //output = "done";
                    
                }
            }
            catch (Exception ex)
            {
                output = ex.Message;
            }
            return output;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if ((_cmd != null) &&
                (_cmd.HasExited != true))
            {
                _cmd.CancelErrorRead();
                _cmd.CancelOutputRead();
                _cmd.Close();
                _cmd.WaitForExit();
            }
        }

        private void _cmd_OutputDataReceived(object sender, DataReceivedEventArgs outLine)
        {
            if (!String.IsNullOrEmpty(outLine.Data))
            {
                //sortOutput.Append(Environment.NewLine + outLine.Data);
                UpdateConsole(Environment.NewLine + outLine.Data);
            }
            
        }

        private void _cmd_ErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            UpdateConsole(e.Data, Brushes.Red);
        }

        private void _cmd_Exited(object sender, EventArgs e)
        {
            _cmd.OutputDataReceived -= new DataReceivedEventHandler(_cmd_OutputDataReceived);
            _cmd.Exited -= new EventHandler(_cmd_Exited);

        }

        private void ScrollViewer_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            outputViewer.ScrollToBottom();
        }

        private void UpdateConsole(string text)
        {
            UpdateConsole(text, null);
        }

        private void UpdateConsole(string text, Brush color)
        {
            if (!output.Dispatcher.CheckAccess())
            {
                output.Dispatcher.Invoke(
                        new Action(
                                () =>
                                {
                                    WriteLine(text, color);
                                }
                            )
                    );
            }
            else
            {
                WriteLine(text, color);
            }
        }

        private void WriteLine(string text, Brush color)
        {
            if (text != null)
            {
                //Span line = new Span();
                //if (color != null)
                //{
                //    line.Foreground = color;
                //}
                //foreach (string textLine in text.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                //{
                //    line.Inlines.Add(new Run(textLine));
                //}
                //line.Inlines.Add(new LineBreak());
                //output.Inlines.Add(line);
                if (text.Length >3)
                {
                    string cleaned = text.Replace("\0", "");
                    string log = getLineBySubstring(cleaned, "===");
                    if (log.StartsWith("==="))
                    {
                        output.AppendText(Environment.NewLine + log.Replace("=", ""));
                    }
                    
                }

            }
        }

        static string getLineBySubstring(String myInput, String mySubstring)
        {
            string[] lines = myInput.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
                if (line.StartsWith(mySubstring))
                    return line;
            return "NaN";
        }
        #endregion

    }
}
