using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

class Installer : Form
{
    private Label lblTitle, lblDesc, lblPath, lblStatus;
    private TextBox txtPath;
    private Button btnBrowse, btnInstall, btnCancel;
    private ProgressBar progress;
    private CheckBox chkDesktop, chkStart;

    [STAThread]
    static void Main() { Application.EnableVisualStyles(); Application.Run(new Installer()); }

    public Installer()
    {
        this.Text = "ShopManager Setup";
        this.Size = new Size(520, 420);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.BackColor = Color.White;

        lblTitle = new Label();
        lblTitle.Text = "ShopManager Installer";
        lblTitle.Font = new Font("Segoe UI", 18, FontStyle.Bold);
        lblTitle.ForeColor = Color.FromArgb(0, 122, 255);
        lblTitle.Location = new Point(30, 25);
        lblTitle.Size = new Size(460, 40);
        this.Controls.Add(lblTitle);

        lblDesc = new Label();
        lblDesc.Text = "Welcome! Click Install to begin installation.";
        lblDesc.Font = new Font("Segoe UI", 10);
        lblDesc.Location = new Point(30, 75);
        lblDesc.Size = new Size(460, 30);
        this.Controls.Add(lblDesc);

        lblPath = new Label();
        lblPath.Text = "Install to:";
        lblPath.Font = new Font("Segoe UI", 10);
        lblPath.Location = new Point(30, 130);
        lblPath.Size = new Size(80, 30);
        this.Controls.Add(lblPath);

        txtPath = new TextBox();
        txtPath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "ShopManager");
        txtPath.Font = new Font("Segoe UI", 10);
        txtPath.Location = new Point(120, 130);
        txtPath.Size = new Size(280, 30);
        this.Controls.Add(txtPath);

        btnBrowse = new Button();
        btnBrowse.Text = "...";
        btnBrowse.Font = new Font("Segoe UI", 10);
        btnBrowse.Location = new Point(410, 129);
        btnBrowse.Size = new Size(60, 30);
        btnBrowse.Click += (s, e) => {
            using (var f = new FolderBrowserDialog())
            {
                f.Description = "Select install folder";
                if (f.ShowDialog() == DialogResult.OK)
                    txtPath.Text = Path.Combine(f.SelectedPath, "ShopManager");
            }
        };
        this.Controls.Add(btnBrowse);

        chkDesktop = new CheckBox();
        chkDesktop.Text = "Create desktop shortcut";
        chkDesktop.Font = new Font("Segoe UI", 10);
        chkDesktop.Location = new Point(120, 180);
        chkDesktop.Size = new Size(250, 30);
        chkDesktop.Checked = true;
        this.Controls.Add(chkDesktop);

        chkStart = new CheckBox();
        chkStart.Text = "Create start menu entry";
        chkStart.Font = new Font("Segoe UI", 10);
        chkStart.Location = new Point(120, 215);
        chkStart.Size = new Size(250, 30);
        chkStart.Checked = true;
        this.Controls.Add(chkStart);

        lblStatus = new Label();
        lblStatus.Text = "";
        lblStatus.Font = new Font("Segoe UI", 9);
        lblStatus.ForeColor = Color.Green;
        lblStatus.Location = new Point(30, 265);
        lblStatus.Size = new Size(460, 25);
        this.Controls.Add(lblStatus);

        progress = new ProgressBar();
        progress.Location = new Point(30, 295);
        progress.Size = new Size(440, 22);
        this.Controls.Add(progress);

        btnInstall = new Button();
        btnInstall.Text = "Install";
        btnInstall.Font = new Font("Segoe UI", 12, FontStyle.Bold);
        btnInstall.BackColor = Color.FromArgb(0, 122, 255);
        btnInstall.ForeColor = Color.White;
        btnInstall.FlatStyle = FlatStyle.Flat;
        btnInstall.Location = new Point(150, 335);
        btnInstall.Size = new Size(120, 42);
        btnInstall.Click += BtnInstall_Click;
        this.Controls.Add(btnInstall);

        btnCancel = new Button();
        btnCancel.Text = "Cancel";
        btnCancel.Font = new Font("Segoe UI", 10);
        btnCancel.Location = new Point(300, 335);
        btnCancel.Size = new Size(100, 42);
        btnCancel.Click += (s, e) => this.Close();
        this.Controls.Add(btnCancel);
    }

    private void BtnInstall_Click(object sender, EventArgs e)
    {
        string installDir = txtPath.Text.Trim();
        if (string.IsNullOrEmpty(installDir))
        {
            MessageBox.Show("Please select an install directory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        btnInstall.Enabled = false;
        btnBrowse.Enabled = false;
        txtPath.Enabled = false;

        try
        {
            progress.Value = 10;
            lblStatus.Text = "Creating directory...";
            Directory.CreateDirectory(installDir);

            progress.Value = 30;
            lblStatus.Text = "Copying files...";
            string sourceDir = Path.GetDirectoryName(Application.ExecutablePath);
            CopyDirectory(sourceDir, installDir);

            progress.Value = 60;
            lblStatus.Text = "Creating shortcuts...";

            if (chkDesktop.Checked)
            {
                string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ShopManager.lnk");
                CreateShortcut(desktopPath, Path.Combine(installDir, "ShopManager.exe"), installDir);
            }

            if (chkStart.Checked)
            {
                string startMenuDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    @"Microsoft\Windows\Start Menu\Programs\ShopManager");
                Directory.CreateDirectory(startMenuDir);
                string shortcutPath = Path.Combine(startMenuDir, "ShopManager.lnk");
                CreateShortcut(shortcutPath, Path.Combine(installDir, "ShopManager.exe"), installDir);
            }

            progress.Value = 90;
            lblStatus.Text = "Finalizing...";

            progress.Value = 100;
            lblStatus.Text = "Installation complete!";

            var result = MessageBox.Show(
                "Installation successful!\n\nLocation: " + installDir + "\n\nStart now?",
                "Done", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Path.Combine(installDir, "ShopManager.exe"),
                    WorkingDirectory = installDir
                });
                System.Threading.Thread.Sleep(3000);
                Process.Start("http://localhost:8000");
            }
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Installation failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            btnInstall.Enabled = true;
            btnBrowse.Enabled = true;
            txtPath.Enabled = true;
        }
    }

    private void CopyDirectory(string source, string dest)
    {
        Directory.CreateDirectory(dest);
        foreach (string file in Directory.GetFiles(source))
        {
            string destFile = Path.Combine(dest, Path.GetFileName(file));
            File.Copy(file, destFile, true);
        }
        foreach (string dir in Directory.GetDirectories(source))
        {
            string destDir = Path.Combine(dest, Path.GetFileName(dir));
            CopyDirectory(dir, destDir);
        }
    }

    private void CreateShortcut(string shortcutPath, string targetPath, string workingDir)
    {
        using (var shell = new Process())
        {
            shell.StartInfo.FileName = "cscript";
            string vbs = string.Format(
                "Set oWS = WScript.CreateObject(\"WScript.Shell\")\n" +
                "Set oLink = oWS.CreateShortcut(\"{0}\")\n" +
                "oLink.TargetPath = \"{1}\"\n" +
                "oLink.WorkingDirectory = \"{2}\"\n" +
                "oLink.Description = \"ShopManager\"\n" +
                "oLink.Save",
                shortcutPath.Replace("\\", "\\\\"),
                targetPath.Replace("\\", "\\\\"),
                workingDir.Replace("\\", "\\\\"));
            string vbsFile = Path.Combine(Path.GetTempPath(), "shortcut.vbs");
            File.WriteAllText(vbsFile, vbs);
            shell.StartInfo.Arguments = "/nologo \"" + vbsFile + "\"";
            shell.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            shell.Start();
            shell.WaitForExit();
            File.Delete(vbsFile);
        }
    }
}
