using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ReleaseAutomation1
{
    public partial class Form1 : Form
    {
        public Dictionary<int, string> listofNames = new Dictionary<int, string>(5);
        public string savePath { get; set; }

        public Form1()
        {
            InitializeComponent();

            button1.Enabled = false;
            label1.ForeColor = Color.Gray;
 
            listofNames.Add(1, "Get files from GIT");
            listofNames.Add(2,"Release on UAT");
            listofNames.Add(3,"Release on TRN");
            listofNames.Add(4,"Release on PROD");
            listofNames.Add(5,"Push on GIT(UAT/TRN)");
            listofNames.Add(6,"Push on GIT(PROD)");

            foreach (KeyValuePair<int, string> keyValue in listofNames)
            {
                comboBox1.Items.Add(keyValue.Value);
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string stringfilename = openFileDialog1.FileName;

                //==============вставляем выбор папки для сохранения после нажатия кнопки ОК===============

                FolderBrowserDialog brwsr = new FolderBrowserDialog();

                //Check to see if the user clicked the cancel button
                if (brwsr.ShowDialog() == DialogResult.Cancel)
                    return;
                else
                {
                    savePath = brwsr.SelectedPath;
                    //Do whatever with the new path
                }
                
                //=========================================================================================
                
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(stringfilename);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                dataGridView1.ColumnCount = colCount;
                dataGridView1.RowCount = rowCount;

                GetFromGit getFiles1 = new GetFromGit();

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                        } 
                    }
                    string filepath123 = getFiles1.chooseFileToDownload(@"C:\FERRERO-RU", xlRange.Cells[i, 3].Value, xlRange.Cells[i, 2].Value);
                    getFiles1.gitDownload(@"C:\FERRERO-RU", @"C:\Development\zipfile.zip", xlRange.Cells[i, 3].Value, filepath123, savePath);
                }

                MessageBox.Show("В xl файле - " + rowCount.ToString() + " рядов. Загружено - " + getFiles1.getCountOfFiles().ToString() + " файлов.");

                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }         
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int key = 0;
            foreach (KeyValuePair<int, string> keyValue in listofNames)
            {
                if (keyValue.Value != comboBox1.SelectedItem.ToString())
                {
                    continue;
                }
                key = keyValue.Key;
                break;               
            }

            if (key == 1)
            {
                button1.Enabled = true;
                label1.ForeColor = Color.Black;
            }
            else
            {
                button1.Enabled = false;
                label1.ForeColor = Color.Gray;
            }
        }

        public string GitCommit(string clientpath, string notes)
        {
            string _commitbat = "git commit -a -m \"" + notes + "\"";
            File.WriteAllText(clientpath + @"\__COMMIT.bat", _commitbat);
            System.Diagnostics.Process process_git = new System.Diagnostics.Process();//Create new process
            System.Diagnostics.ProcessStartInfo startInfo_git = new System.Diagnostics.ProcessStartInfo//Add start info for process
            {
                UseShellExecute = false, //Use shell commands (NO GUI flag)
                RedirectStandardOutput = true, //Need to return output
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, //Hide CMD window
                WorkingDirectory = clientpath, //Path for Pull or Push
                FileName = clientpath + @"\__COMMIT.bat"//path to CMD
            };
            process_git.StartInfo = startInfo_git;//Add params for process and start
            process_git.Start();
            string output = process_git.StandardOutput.ReadToEnd();
            process_git.WaitForExit();
            File.Delete(clientpath + @"\__COMMIT.bat");
            string[] outputarr = output.Split('\n');
            return outputarr[1]+Environment.NewLine+ outputarr[3];
        }
    }
}
