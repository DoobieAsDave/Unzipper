using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using System.IO.Compression;
using System.Security.Principal;

namespace Unzipper
{
    public enum AlertIcon
    {
        Success,
        Error,
        Info,
        Quest
    }

    public partial class Main : Form
    {
        private string selectedDirectoryPath, destinationPath;

        private List<string> detectedZipFilePaths;
        private List<BackgroundWorker> backgroundWorkers = null;

        private bool overwrite, unzipping, unzippingCancelled = false;
        private int unzipCounter = 0;        

        public Main()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            try
            {
                if (CheckPermission())
                {
                    SetDirectoryPath(GetLastUsedDirectory());
                }
                else
                {
                    var alert = new Alert("Not enuff Permission", "You need to run the Programm as an Administrator!", AlertIcon.Info, null, null);
                    alert.ShowDialog();

                    Close();
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private bool CheckPermission()
        {
            try
            {
                return (new WindowsPrincipal(WindowsIdentity.GetCurrent()))
                    .IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();

                return false;
            }
        }
        private string GetLastUsedDirectory()
        {
            try
            {
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "LastDirectory.txt"))
                {
                    string lastDirectory = (File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + @"LastDirectory.txt") != string.Empty
                                                ? File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + @"LastDirectory.txt")
                                                : @"C:\Users\" + Environment.UserName + @"\Downloads");

                    if (Directory.Exists(lastDirectory))
                    {
                        return lastDirectory;
                    }
                    else
                    {
                        return @"C:\Users\" + Environment.UserName + @"\Downloads";
                    }                    
                }
                else
                {
                    return @"C:\Users\" + Environment.UserName + @"\Downloads";
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();

                return @"C:\Users\" + Environment.UserName + @"\Downloads";
            }
        }
        
        private void btnSelectDirectory_Click(object sender, EventArgs e)
        {
            try
            {
                var directoryDialog = new FolderBrowserDialog();
                directoryDialog.SelectedPath = @"C:\Users\" + Environment.UserName + @"Downloads";
                directoryDialog.ShowNewFolderButton = false;
                directoryDialog.Description = "Select the Directory which contains the Zip Files";

                if (directoryDialog.ShowDialog() == DialogResult.OK)
                {
                    SetDirectoryPath(directoryDialog.SelectedPath);                    
                }                
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private void btnDirectoryRescan_Click(object sender, EventArgs e)
        {
            try
            {
                SetDirectoryPath(selectedDirectoryPath);
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void SetDirectoryPath(string directoryPath)
        {
            try
            {
                selectedDirectoryPath = directoryPath;
                lblShowDirectory.Text = directoryPath;

                if (selectedDirectoryPath != @"C:\Users\" + Environment.UserName + @"\Downloads")
                {
                    lblShowDirectory.ForeColor = Color.Green;
                }
                else
                {
                    lblShowDirectory.ForeColor = Color.Coral;
                }

                SetFoundFilesToList(
                        detectedZipFilePaths = ScanSelectedDirectory()
                    );
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private List<string> ScanSelectedDirectory()
        {
            try
            {
                return Directory.GetFiles(selectedDirectoryPath, "*.zip").ToList();                
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();

                return null;
            }
        }
        private void SetFoundFilesToList(List<string> filePaths)
        {
            try
            {
                if (filePaths != null)
                {
                    lsbDetectedZipFiles.Items.Clear();           
                    lsbDetectedZipFiles.Items.AddRange(filePaths.Select(p => Path.GetFileNameWithoutExtension(p)).ToArray());

                    btnUnzipSelectedFiles.Enabled = true;
                    btnUnzipSelectedFiles.Text = "Unzip Files";
                    btnUnzipSelectedFiles.BackColor = Color.LightGreen;
                }
                else
                {
                    var alert = new Alert("No Zip Files found", "There are no Zip Files to decompress in the selected Directory:\n" + selectedDirectoryPath + "\n\n Please choose a other Directory to start a decompression...", AlertIcon.Info, null, null);
                    alert.ShowDialog();

                    btnUnzipSelectedFiles.Enabled = false;
                    btnUnzipSelectedFiles.Text = "Unzip Files";
                    btnUnzipSelectedFiles.BackColor = Color.LightCoral;
                }                
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private void lsbDetectedZipFiles_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {                
                if (!unzipping)
                {
                    if (detectedZipFilePaths.Count > 0)
                    {
                        if (lsbDetectedZipFiles.SelectedItem != null)
                        {
                            var alert = new Alert("Remove Zip from List", "Do you want to remove following Zip File from List?\n\n" + lsbDetectedZipFiles.SelectedItem.ToString(), AlertIcon.Quest, null, null);
                            alert.ShowDialog();

                            if (alert.Return)
                            {
                                detectedZipFilePaths.RemoveAll(p => p.Contains(lsbDetectedZipFiles.SelectedItem.ToString()));

                                SetFoundFilesToList(detectedZipFilePaths);
                            }
                        }
                        else
                        {
                            var alert = new Alert("No Zip File selected", "Please double-click a Zip File from the List to remove it...", AlertIcon.Info, null, null);
                            alert.ShowDialog();
                        }
                    }
                    else
                    {
                        var alert = new Alert("All Zip Files removed", "You have removed all the detected Zip Files!\n\nPlease select a Directory to start unzipping...", AlertIcon.Info, null, null);
                        alert.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void btnUnzipSelectedFiles_Click(object sender, EventArgs e)
        {
            try
            {                
                if (!unzipping)
                {                                      
                    var folderDialog = new FolderBrowserDialog();
                    folderDialog.ShowNewFolderButton = false;
                    folderDialog.SelectedPath = selectedDirectoryPath;
                    folderDialog.Description = "Select the Destination Folder for the unzipped Files";

                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        destinationPath = folderDialog.SelectedPath;

                        unzipping = true;

                        btnSelectDirectory.Enabled = false;
                        btnDirectoryRescan.Enabled = false;
                        pgbUnzippingStatus.Visible = true;

                        btnUnzipSelectedFiles.Text = "Stop unzipping";
                        btnUnzipSelectedFiles.BackColor = Color.LightCoral;

                        if (CheckIfDirectoriesAlreadyExists())
                        {
                            if (backgroundWorkers == null)
                            {
                                backgroundWorkers = new List<BackgroundWorker>();
                            }

                            foreach (var zipPath in detectedZipFilePaths)
                            {
                                var backgroundWorker = new BackgroundWorker();
                                backgroundWorker.WorkerSupportsCancellation = true;
                                backgroundWorker.DoWork += (se, ev) => backgroundWorker_DoWork(se, ev, zipPath);
                                backgroundWorker.RunWorkerCompleted += (se, ev) => backgroundWorker_RunWorkerCompleted(se, ev, Path.GetFileNameWithoutExtension(zipPath));
                                backgroundWorker.RunWorkerAsync();
                                backgroundWorkers.Add(backgroundWorker);
                            }
                        }
                        else
                        {
                            unzipping = false;

                            btnUnzipSelectedFiles.Text = "Unzip Files";
                            btnUnzipSelectedFiles.BackColor = Color.LightGreen;

                            btnSelectDirectory.Enabled = true;
                            btnDirectoryRescan.Enabled = true;
                            pgbUnzippingStatus.Visible = false;
                            pgbUnzippingStatus.Value = 0;

                            if (overwrite)
                            {
                                var alert = new Alert("No Files to unzip", "You exluded all the Zip Files!\n\nSo there is nothing to unzip...", AlertIcon.Success, null, null);
                                alert.ShowDialog();
                            }
                            else
                            {
                                var alert = new Alert("No Files to unzip", "You choosed to not overwrite the existing Directories!\n\nSo there is nothing to unzip...", AlertIcon.Success, null, null);
                                alert.ShowDialog();
                            }
                        }
                    }                                      
                }
                else
                {
                    Cursor = Cursors.WaitCursor;                    

                    unzipping = false;
                    unzippingCancelled = true;

                    btnUnzipSelectedFiles.Text = "Canelling Unzipping";
                    btnUnzipSelectedFiles.BackColor = Color.LightGreen;

                    pgbUnzippingStatus.Visible = false;
                    pgbUnzippingStatus.Value = 0;

                    foreach (var backgroundWorker in backgroundWorkers)
                    {
                        backgroundWorker.CancelAsync();
                    }                    

                    System.Threading.Thread.Sleep(3000);

                    RemoveUnzippedFiles();

                    var alert = new Alert("Unzipping stopped", "The Unzipping has been stopped and all the unzipped Directories have been deleted...", AlertIcon.Info, null, null);
                    alert.ShowDialog();

                    btnUnzipSelectedFiles.Text = "Unzip Files";

                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private void RemoveUnzippedFiles()
        {
            try
            {
                string[] deletationPaths = detectedZipFilePaths.Select(p => destinationPath + @"\" + Path.GetFileNameWithoutExtension(p)).ToArray();

                foreach (var deletationPath in deletationPaths)
                {
                    if (Directory.Exists(deletationPath))
                    {
                        DeleteDataFromDirectory(deletationPath);
                        Directory.Delete(deletationPath);
                    }
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private bool CheckIfDirectoriesAlreadyExists()
        {
            try
            {
                List<string> existingDirectoriesPaths = new List<string>();

                foreach (var zipPath in detectedZipFilePaths)
                {
                    if (Directory.Exists(destinationPath + @"\" + Path.GetFileNameWithoutExtension(zipPath)))
                    {
                        existingDirectoriesPaths.Add(zipPath);
                    }
                }

                detectedZipFilePaths.RemoveAll(p => existingDirectoriesPaths.Contains(p));

                if (existingDirectoriesPaths.Count != 0)
                {
                    string alertMessage = "The following Directories already exists in your Destination Directory!\n\nDo you want to overwrite them?\n\nYou can also exclude Directories from overwriting by selecting them";
                    var alert = new Alert("Overwrite existing Directories", alertMessage, AlertIcon.Quest, null, existingDirectoriesPaths);
                    alert.ShowDialog();

                    if (alert.Return)
                    {
                        overwrite = true;

                        if (alert.OverwritePaths != null)
                        {
                            detectedZipFilePaths.AddRange(alert.OverwritePaths);

                            SetFoundFilesToList(detectedZipFilePaths.ToList());
                        }
                    }
                    else
                    {
                        overwrite = false;
                    }
                }
                else
                {
                    overwrite = false;
                }

                if (detectedZipFilePaths.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();

                return false;
            }
        }
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e, string zipPath)
        {
            try
            {
                if (!((BackgroundWorker)sender).CancellationPending)
                {
                    string directoryPath = destinationPath + @"\" + Path.GetFileNameWithoutExtension(zipPath);

                    if (Directory.Exists(directoryPath))
                    {
                        if (overwrite)
                        {
                            DeleteDataFromDirectory(directoryPath);
                            Directory.Delete(directoryPath);
                            Directory.CreateDirectory(directoryPath);

                            ZipFile.ExtractToDirectory(zipPath, directoryPath);

                            unzipCounter += 1;
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(directoryPath);

                        ZipFile.ExtractToDirectory(zipPath, directoryPath);

                        unzipCounter += 1;
                    }
                }
                else
                {
                    e.Cancel = true;
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e, string zipFilename)
        {
            try
            {
                if (!e.Cancelled)
                {
                    bool afterScan = true;

                    foreach (var zipFile in lsbDetectedZipFiles.Items)
                    {
                        if (zipFile.ToString() == zipFilename)
                        {
                            lsbDetectedZipFiles.Items.Remove(zipFile);
                            break;
                        }
                    }

                    UpdateProgressBar();

                    if (backgroundWorkers.Find(b => b.IsBusy == true) == null)
                    {                       
                        if (!unzippingCancelled)
                        {
                            btnUnzipSelectedFiles.Text = "Unzip Files";
                            btnUnzipSelectedFiles.BackColor = Color.LightGreen;
                            pgbUnzippingStatus.Visible = false;
                            pgbUnzippingStatus.Value = 0;

                            if (unzipCounter != 0)
                            {
                                lsbDetectedZipFiles.Items.Clear();                                

                                var alert = new Alert("Unzipping completed successfuly", "The " + unzipCounter + " Zip Files have been successfuly unzipped to:\n" + destinationPath + "\n\nWould you like to delete the Zip Files?", AlertIcon.Success, destinationPath, null);
                                alert.ShowDialog();

                                if (alert.Return)
                                {
                                    afterScan = false;
                                    DeleteZipFiles();
                                }
                            }
                            else
                            {
                                var alert = new Alert("No Files to unzip", "You choosed to not overwrite the existing Directories!\n\nSo there is nothing to unzip...", AlertIcon.Success, null, null);
                                alert.ShowDialog();
                            }
                        }

                        btnSelectDirectory.Enabled = true;
                        btnDirectoryRescan.Enabled = true;

                        backgroundWorkers = null;

                        unzipping = false;
                        overwrite = false;
                        unzipCounter = 0;                        

                        if (afterScan)
                        {
                            SetDirectoryPath(selectedDirectoryPath);
                        }                        
                    }
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private void DeleteDataFromDirectory(string directoryPath)
        {
            try
            {
                var directoryInfo = new DirectoryInfo(directoryPath);

                foreach (var file in directoryInfo.GetFiles())
                {
                    file.Delete();
                }

                foreach (var directory in directoryInfo.GetDirectories())
                {
                    directory.Delete();
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }        
        private void DeleteZipFiles()
        {
            try
            {
                foreach (string zipPath in detectedZipFilePaths)
                {
                    File.Delete(zipPath);
                }

                btnUnzipSelectedFiles.Enabled = false;
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }
        private void UpdateProgressBar()
        {
            try
            {
                if (!unzippingCancelled)
                {
                    if (pgbUnzippingStatus.Value <= 100)
                    {
                        pgbUnzippingStatus.Value += 100 / detectedZipFilePaths.Count;
                    }
                    else
                    {
                        pgbUnzippingStatus.Value = 100;
                    }
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (backgroundWorkers != null)
                {
                    foreach (var backgroundWorker in backgroundWorkers)
                    {
                        if (backgroundWorker.IsBusy)
                        {
                            var alert = new Alert("Unzipping running", "A Unzipping Process is still running!\n\nDo you really want to close the Programm?\nIt could manipulate the Files!", AlertIcon.Quest, null, null);
                            alert.ShowDialog();

                            if (alert.Return)
                            {
                                foreach (var worker in backgroundWorkers)
                                {
                                    worker.CancelAsync();
                                }

                                break;
                            }
                            else
                            {
                                e.Cancel = true;
                            }
                        }
                    }
                }

                if (selectedDirectoryPath != null)
                {
                    if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + @"LastDirectory.txt"))
                    {
                        File.Create(AppDomain.CurrentDomain.BaseDirectory + @"LastDirectory.txt");
                    }

                    File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + @"LastDirectory.txt", string.Empty);
                    File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + @"LastDirectory.txt", selectedDirectoryPath);
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }        
    }    
}
