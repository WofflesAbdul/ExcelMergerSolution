using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

public partial class Form1 : Form, IFileSelectionView
{
    private const string PlaceholderDuringCreateNewMode = "Enter filename here";
    private readonly FileSelectionPresenter presenter;
    private readonly List<Control> controlsToDisable = new List<Control>();
    private readonly List<ToolStripMenuItem> toolStripButtonsToDisable = new List<ToolStripMenuItem>();
    private Dictionary<Control, bool> controlsStateBackup;
    private Dictionary<ToolStripMenuItem, bool> toolStripStateBackup;
    private CancellationTokenSource progressAnimationCts;

    public event EventHandler MergeClicked;

    public event EventHandler SortClicked;

    public event EventHandler ResetClicked;

    public event EventHandler OpenFileClicked;

    public event EventHandler OpenFolderClicked;

    public event EventHandler<BaseFileModeSelection> BaseFileModeChanged;

    public Form1()
    {
        InitializeComponent();
        presenter = new FileSelectionPresenter(this);

        controlsToDisable.AddRange(new Control[] { button1, button2, button3, button4, });
        toolStripButtonsToDisable.AddRange(new ToolStripMenuItem[] { openFileToolStripMenuItem, openContainingFolderToolStripMenuItem, resetToolStripMenuItem, });

        // Subscribe to requested events
        MergeClicked += (s, e) => presenter.OnMergeClicked();
        SortClicked += (s, e) => presenter.OnSortClicked();
        ResetClicked += (s, e) => presenter.OnResetClicked();
        presenter.MergeRequested += async (s, e) => await RunMergeAsync();
        presenter.SortRequested += async (s, e) => await RunSortAsync();
        presenter.ResetRequested += (s, e) => presenter.ResetSelection();
        rbUseExistingFile.Checked = true;
    }

    public BaseFileModeSelection CurrentBaseFileMode => rbUseExistingFile.Checked ? BaseFileModeSelection.UseExistingFile : rbCreateNewFile.Checked ? BaseFileModeSelection.CreateNewFile : throw new InvalidOperationException();

    public void UpdateBaseFileName(string name)
    {
        textBox1.Text = name;
        presenter.CheckMergeButtonState();
        presenter.CheckSortButtonState();
    }

    public void UpdateBaseFileFolderName(string name)
    {
        label4.Text = name;
    }

    public void UpdateTargetFileNames(string names)
    {
        textBox2.Text = names;
        presenter.CheckMergeButtonState();
    }

    public void SetMergeButtonEnabled(bool enabled)
    {
        button3.Enabled = enabled;
    }

    public void SetSortButtonEnabled(bool enabled)
    {
        button4.Enabled = enabled;
    }

    public void SetOpenFileButtonEnabled(bool enabled)
    {
        openFileToolStripMenuItem.Enabled = enabled;
        openContainingFolderToolStripMenuItem.Enabled = enabled;
    }

    public void ApplyTaskControlLock(bool disableForTask)
    {
        if (!disableForTask)
        {
            // Initialize backup dictionaries
            controlsStateBackup = new Dictionary<Control, bool>();
            foreach (var ctrl in controlsToDisable)
            {
                controlsStateBackup[ctrl] = ctrl.Enabled; // store current state
                ctrl.Enabled = false; // disable
            }

            toolStripStateBackup = new Dictionary<ToolStripMenuItem, bool>();
            foreach (var btn in toolStripButtonsToDisable)
            {
                toolStripStateBackup[btn] = btn.Enabled;
                btn.Enabled = false;
            }
        }
        else
        {
            // Restore from backup
            if (controlsStateBackup != null)
            {
                foreach (var kvp in controlsStateBackup)
                {
                    kvp.Key.Enabled = kvp.Value;
                }
            }

            if (toolStripStateBackup != null)
            {
                foreach (var kvp in toolStripStateBackup)
                {
                    kvp.Key.Enabled = kvp.Value;
                }
            }

            // Clear backups
            controlsStateBackup?.Clear();
            toolStripStateBackup?.Clear();
        }
    }

    public void SetProgress(int percent)
    {
        if (percent < toolStripProgressBar1.Minimum)
        {
            percent = toolStripProgressBar1.Minimum;
        }

        if (percent > toolStripProgressBar1.Maximum)
        {
            percent = toolStripProgressBar1.Maximum;
        }

        toolStripProgressBar1.Value = percent;
    }

    public DialogResult ShowPrompt(string message, string title)
    {
        return MessageBox.Show(
            message,
            title,
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        );
    }

    public void UpdateUIForBaseFileSelectionMode()
    {
        textBox1.Text = string.Empty;
        label4.Text = string.Empty;
        switch (CurrentBaseFileMode)
        {
            case BaseFileModeSelection.UseExistingFile:
                textBox1.ReadOnly = true;          // cannot edit filename
                button1.Text = "Select File";
                break;

            case BaseFileModeSelection.CreateNewFile:
                textBox1.ReadOnly = false;         // user can type filename
                TextBox1_Leave(textBox1, EventArgs.Empty);
                button1.Text = "Select Folder";
                break;
        }
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        toolStripLabel2.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        toolStripLabel2.AutoSize = true;
    }

    private void Button1_Click(object sender, EventArgs e)
    {
        switch (CurrentBaseFileMode)
        {
            case BaseFileModeSelection.UseExistingFile:
                presenter.SelectBaseFile();
                break;

            case BaseFileModeSelection.CreateNewFile:
                presenter.OnNewBaseFileTargetLocationChanged();
                break;
        }
    }

    private void Button2_Click(object sender, EventArgs e)
    {
        presenter.SelectTargetFiles();
    }

    private void Button3_Click(object sender, EventArgs e)
    {
        MergeClicked?.Invoke(this, EventArgs.Empty);
    }

    private void Button4_Click(object sender, EventArgs e)
    {
        SortClicked?.Invoke(this, EventArgs.Empty);
    }

    private void ButtonReset_Click(object sender, EventArgs e)
    {
        ResetClicked?.Invoke(this, EventArgs.Empty);
    }

    private void ButtonOpenFile_Click(object sender, EventArgs e)
    {
        OpenFileClicked?.Invoke(this, EventArgs.Empty);
    }

    private void ButtonOpenFolder_Click(object sender, EventArgs e)
    {
        OpenFolderClicked?.Invoke(this, EventArgs.Empty);
    }

    private void RbUseExistingFile_CheckedChanged(object sender, EventArgs e)
    {
        if (rbUseExistingFile.Checked)
        {
            BaseFileModeChanged?.Invoke(this, BaseFileModeSelection.UseExistingFile);
        }
    }

    private void RbCreateNewFile_CheckedChanged(object sender, EventArgs e)
    {
        if (rbCreateNewFile.Checked)
        {
            BaseFileModeChanged?.Invoke(this, BaseFileModeSelection.CreateNewFile);
        }
    }

    private void TextBox1_Enter(object sender, EventArgs e)
    {
        if (CurrentBaseFileMode == BaseFileModeSelection.CreateNewFile &&
            textBox1.Text == PlaceholderDuringCreateNewMode)
        {
            // Clear placeholder
            textBox1.Text = string.Empty;
            textBox1.ForeColor = SystemColors.WindowText;
            textBox1.Font = new Font(textBox1.Font, FontStyle.Regular);

            // Clear the model value as user starts typing
        }
    }

    private void TextBox1_Leave(object sender, EventArgs e)
    {
        if (CurrentBaseFileMode == BaseFileModeSelection.CreateNewFile)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                // Restore placeholder if empty
                textBox1.Text = PlaceholderDuringCreateNewMode;
                textBox1.ForeColor = SystemColors.GrayText;
                textBox1.Font = new Font(textBox1.Font, FontStyle.Italic);

                presenter.OnNewBaseFilenameEntered(null);
            }
            else if (textBox1.Text != PlaceholderDuringCreateNewMode)
            {
                // Save user input in model
                presenter.OnNewBaseFilenameEntered(textBox1.Text.Trim());
            }
        }
    }

    private void TextBox1_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && CurrentBaseFileMode == BaseFileModeSelection.CreateNewFile)
        {
            e.SuppressKeyPress = true; // Prevent newline
            this.SelectNextControl(textBox1, forward: true, tabStopOnly: true, nested: true, wrap: true);
        }
    }

    private async Task RunMergeAsync()
    {
        ApplyTaskControlLock(false);
        presenter.SetProgress(0); // reset

        if (CurrentBaseFileMode == BaseFileModeSelection.CreateNewFile)
        {
            string tempBaseFilePath = presenter.CreateNewBaseFile();

            rbUseExistingFile.Checked = true; // now new file exist, hence use existing, this resets some model property and UI
            presenter.SelectBaseFile(tempBaseFilePath); // reassigned model property and UI
        }

        try
        {
            var merger = new Excel_Handling.ExcelMerger();
            await Task.Run(() =>
            {
                merger.MergeFiles(
                    presenter.BaseFilePath,
                    presenter.TargetFilePaths,
                    percent => this.Invoke((Action)(() => SetProgress(percent)))
                );
            });

            // If the base file was created as a blank workbook, remove the placeholder sheet
            if (CurrentBaseFileMode == BaseFileModeSelection.CreateNewFile)
            {
                presenter.RemovePlaceholderIfNeeded();
            }

            presenter.NotifyMergeCompleted();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Merge failed: {ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            ApplyTaskControlLock(true);
            presenter.ClearTargetSelection();
            presenter.SetProgress(0); // reset
        }
    }

    private async Task RunSortAsync()
    {
        ApplyTaskControlLock(false);

        try
        {
            // Reset and start animation
            toolStripProgressBar1.Value = 0;
            progressAnimationCts = new CancellationTokenSource();
            var animationTask = AnimateProgressBarAsync(toolStripProgressBar1, 20, 200, progressAnimationCts.Token);

            // Run the sort
            var sorter = new Excel_Handling.FunctionalTestSorter();
            await Task.Run(() => sorter.SortSheets(presenter.BaseFilePath)); // Once completed, the vb.net code exits here

            // Cancel the animation once sorting is done
            progressAnimationCts.Cancel();

            try
            {
                await animationTask; // allow clean exit
            }
            catch (TaskCanceledException) { }

            // Snap to 100%
            toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
            presenter.NotifySortCompleted();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sort failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            ApplyTaskControlLock(true);
            toolStripProgressBar1.Value = 0;
        }
    }

    private async Task AnimateProgressBarAsync(ToolStripProgressBar progressBar, int steps, int delayMs, CancellationToken token)
    {
        progressBar.Value = 0;
        int maxValue = (int)(progressBar.Maximum * 0.9); // cap at 90%
        int stepValue = maxValue / steps;

        for (int i = 1; i <= steps; i++)
        {
            if (token.IsCancellationRequested) break;
            await Task.Delay(delayMs, token);
            if (token.IsCancellationRequested) break;
            progressBar.Value = Math.Min(stepValue * i, maxValue);
        }
    }
}
