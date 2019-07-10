using ICSharpCode.SharpZipLib.Zip;
using System.IO;
using System.Windows.Forms;

namespace ReleaseAutomation1
{
    class GetFromGit
    {       
        private int countOfFiles = 0;
        public string s;
        public string gitPull(string clientpath)
        {
            string _commitbat = "git pull --progress -v --no-rebase \"origin\"";
            File.WriteAllText(clientpath + @"\__PULL.bat", _commitbat);
            System.Diagnostics.Process process_git = new System.Diagnostics.Process();//Create new process
            System.Diagnostics.ProcessStartInfo startInfo_git = new System.Diagnostics.ProcessStartInfo//Add start info for process
            {
                UseShellExecute = false, //Use shell commands (NO GUI flag)
                RedirectStandardOutput = true, //Need to return output
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, //Hide CMD window
                WorkingDirectory = clientpath, //Path for Pull or Push
                FileName = clientpath + @"\__PULL.bat"//path to CMD
            };
            process_git.StartInfo = startInfo_git;//Add params for process and start
            process_git.Start();
            string output = process_git.StandardOutput.ReadToEnd();
            process_git.WaitForExit();
            File.Delete(clientpath + @"\__PULL.bat");
            string[] outputarr = output.Split('\n');
            return outputarr[2];
        }

        public void UnzipFile(string zipFileName, string targetDir)
        {
            FastZip fastZip = new FastZip();
            string fileFilter = null;
            // Will always overwrite if target filenames already exist
            fastZip.ExtractZip(zipFileName, targetDir, fileFilter);
        }

        public string chooseFileToDownload(string clientpath, string revision, string fileLookingFor)//передаём локальный путь гита, рефизию и файл который надо искать в ревизии
        {
            string _chooseFilebat = "git show --pretty=" + "\"" + "format:" + "\"" + " --name-only " + revision;
            File.WriteAllText(clientpath + @"\__CHOOSEFILE.bat", _chooseFilebat);
            System.Diagnostics.Process process_git = new System.Diagnostics.Process();//Create new process
            System.Diagnostics.ProcessStartInfo startInfo_git = new System.Diagnostics.ProcessStartInfo//Add start info for process
            {
                UseShellExecute = false, //Use shell commands (NO GUI flag)
                RedirectStandardOutput = true, //Need to return output
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, //Hide CMD window
                WorkingDirectory = clientpath, //Path for Pull or Push
                FileName = clientpath + @"\__CHOOSEFILE.bat"//path to CMD
            };
            process_git.StartInfo = startInfo_git;//Add params for process and start
            process_git.Start();
            string output = process_git.StandardOutput.ReadToEnd();
            process_git.WaitForExit();
            File.Delete(clientpath + @"\__CHOOSEFILE.bat");
            string[] outputarr = output.Split('\n');
            string chosenFileName = "";
            

            for (int i = 2; i <= outputarr.Length-1; i++)
            {
                chosenFileName = Path.GetFileName(outputarr[i]);

                if (chosenFileName.Equals(fileLookingFor))//если файл найден вревизии
                {
                    countOfFiles++;
                    return outputarr[i];
                }
                else
                {                    
                    continue;
                }

            }
            return "File not found";
        }

        public int getCountOfFiles()
        {
            return countOfFiles;
        }

        public string gitDownload(string clientpath, string outputPath, string revision, string fileToDownload, string savePath)
        {
         
            string _commitbat = "git archive --format zip --output " + outputPath + " " + revision + " " + "\"" + fileToDownload + "\"";
            File.WriteAllText(clientpath + @"\__DOWNLOAD.bat", _commitbat);
            System.Diagnostics.Process process_git = new System.Diagnostics.Process();//Create new process
            System.Diagnostics.ProcessStartInfo startInfo_git = new System.Diagnostics.ProcessStartInfo//Add start info for process
            {
                UseShellExecute = false, //Use shell commands (NO GUI flag)
                RedirectStandardOutput = true, //Need to return output
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, //Hide CMD window
                WorkingDirectory = clientpath, //Path for Pull or Push
                FileName = clientpath + @"\__DOWNLOAD.bat"//path to CMD
            };
            process_git.StartInfo = startInfo_git;//Add params for process and start
            process_git.Start();
            string output = process_git.StandardOutput.ReadToEnd();
            process_git.WaitForExit();
            File.Delete(clientpath + @"\__DOWNLOAD.bat");
                        
            UnzipFile(outputPath, savePath);//unzipping file to any dir(2nd argument)
            
            File.Delete(outputPath);//delete zip file

            string[] outputarr = output.Split('\n');
            return outputarr[1];
        }
    }
}
