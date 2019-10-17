using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Unzipper
{
    public partial class Alert : Form
    {
        public bool Return { get; set; } = false;
        public List<string> OverwritePaths { get; set; } = null;

        private string destinationPath;
        private List<string> existingDirectoriesPaths;

        private ListBox listbox;

        public Alert(string alertTitle, string alertMessage, AlertIcon alertIcon, string destinationPath, List<string> existingDirectoriesPaths)
        {
            try
            {
                InitializeComponent();

                if (alertMessage.Contains("\n"))
                {
                    txtAlertMessage.Location = new Point(58, 23);

                    if (existingDirectoriesPaths == null)
                    {
                        txtAlertMessage.Size = new Size(221, 66);
                    }
                    else
                    {
                        Size = new Size(305, 350);

                        txtAlertMessage.Size = new Size(221, 75);

                        listbox = new ListBox();
                        listbox.Location = new Point(58, 100);
                        listbox.Size = new Size(221, 173);
                        listbox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
                        listbox.SelectionMode = SelectionMode.MultiSimple;
                        listbox.Items.AddRange(existingDirectoriesPaths.Select(d => System.IO.Path.GetFileNameWithoutExtension(d)).ToArray());
                        listbox.SelectedIndexChanged += new EventHandler(listbox_SelectedIndexChanged);
                        Controls.Add(listbox);

                        Size = new Size(450, 350);
                    }
                }                

                switch (alertIcon)
                {
                    case AlertIcon.Success:
                        pcbAlertIcon.Image = Properties.Resources.Success;

                        if (alertTitle != "No Files to unzip")
                        {
                            btnReturnOk.Visible = false;
                            btnReturnYes.Visible = true;
                            btnReturnNo.Visible = true;
                        }                                               
                        break;
                    case AlertIcon.Error:
                        pcbAlertIcon.Image = Properties.Resources.Error;
                        break;
                    case AlertIcon.Info:
                        pcbAlertIcon.Image = Properties.Resources.Information;
                        break;
                    case AlertIcon.Quest:
                        pcbAlertIcon.Image = Properties.Resources.Question;

                        btnReturnOk.Visible = false;
                        btnReturnYes.Visible = true;
                        btnReturnNo.Visible = true;
                        break;
                }                

                Text = alertTitle;
                txtAlertMessage.Text = alertMessage;
                this.destinationPath = destinationPath;
                this.existingDirectoriesPaths = existingDirectoriesPaths;

                OverwritePaths = existingDirectoriesPaths;

                if (destinationPath != null)
                {
                    txtAlertMessage.SelectionStart = txtAlertMessage.Text.IndexOf(destinationPath);
                    txtAlertMessage.SelectionLength = destinationPath.Length;
                    txtAlertMessage.SelectionColor = Color.Green;                    
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void txtAlertMessage_Enter(object sender, EventArgs e)
        {
            try
            {
                if (btnReturnOk.Visible)
                {
                    btnReturnOk.Focus();
                }
                else
                {
                    btnReturnYes.Focus();
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void button_Click(object sender, EventArgs e)
        {
            try
            {               
                if (((Button)sender).Text == "Yes")
                {
                    Return = true;
                }

                Close();
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void txtAlertMessage_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (destinationPath != null)
                {
                    System.Diagnostics.Process.Start(destinationPath);
                }
            }
            catch (Exception ex)
            {
                var alert = new Alert("An Error occurred", ex.Message, AlertIcon.Error, null, null);
                alert.ShowDialog();
            }
        }

        private void listbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int[] selectedIndexes = listbox.SelectedIndices.Cast<int>().ToArray();
                List<int> unselectedIndexes = new List<int>();

                for (int i = 0; i < listbox.Items.Count; i++)
                {
                    if (!selectedIndexes.Contains(i))
                    {
                        unselectedIndexes.Add(i);
                    }
                }

                if (unselectedIndexes.Count > 0)
                {
                    OverwritePaths = new List<string>();

                    foreach (var unselectedIndex in unselectedIndexes)
                    {
                        OverwritePaths.Add(existingDirectoriesPaths[unselectedIndex]);
                    }
                }
                else
                {
                    OverwritePaths = null;
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
