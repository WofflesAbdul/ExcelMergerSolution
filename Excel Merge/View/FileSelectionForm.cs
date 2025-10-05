using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

public partial class FileSelectionForm : Form, IFileSelectionView
{
    private const string PlaceholderDuringCreateNewMode = "Enter filename here";
    private readonly FileSelectionPresenter presenter;
    private readonly FileSelectionModel model;
    private readonly List<Control> controlsToDisable = new List<Control>();
    private readonly List<ToolStripMenuItem> toolStripButtonsToDisable = new List<ToolStripMenuItem>();
    private CancellationTokenSource progressAnimationCts;

    public event EventHandler MergeClicked;

    public event EventHandler SortClicked;

    public event EventHandler ResetClicked;

    public event EventHandler OpenFileClicked;

    public event EventHandler OpenFolderClicked;

    public event EventHandler<InputFileMode> InputFileModeChanged;

    public FileSelectionForm()
    {
        InitializeComponent();
        model = new FileSelectionModel();
        presenter = new FileSelectionPresenter(model, this);

        controlsToDisable.AddRange(new Control[] { button1, button2, button3, button4, });
        toolStripButtonsToDisable.AddRange(new ToolStripMenuItem[] { openFileToolStripMenuItem, openContainingFolderToolStripMenuItem, resetToolStripMenuItem, });

        // Subscribe to requested events
        MergeClicked += (s, e) => presenter.OnMergeClicked();
        SortClicked += (s, e) => presenter.OnSortClicked();
        ResetClicked += (s, e) => presenter.OnResetClicked();
        presenter.MergeRequested += async (s, e) => await RunMergeAsync();
        presenter.SortRequested += async (s, e) => await RunSortAsync();
        presenter.ResetRequested += (s, e) => presenter.OnReset();
        rbUseExistingFile.Checked = true;
    }

    public InputFileMode CurrentBaseFileMode =>
        rbUseExistingFile.Checked
            ? InputFileMode.ExistingFile
            : rbCreateNewFile.Checked
                ? InputFileMode.NewFile
                : throw new InvalidOperationException();

    public void UpdateFilename(string name)
    {
        textBox1.Text = name;
    }

    public void UpdateDirectory(string name)
    {
        label4.Text = name;
    }

    public void UpdateTargetFilenames(string names)
    {
        textBox2.Text = names;
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

    public void LockControls(bool enable)
    {
        bool lockEnabled = !enable;
        foreach (var ctrl in controlsToDisable)
        {
            ctrl.Enabled = lockEnabled;
        }

        foreach (var btn in toolStripButtonsToDisable)
        {
            btn.Enabled = lockEnabled;
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
            case InputFileMode.ExistingFile:
                textBox1.ReadOnly = true;
                button1.Text = "Select File";
                break;

            case InputFileMode.NewFile:
                textBox1.ReadOnly = false;
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
            case InputFileMode.ExistingFile:
                presenter.SelectBaseFile();
                break;
            case InputFileMode.NewFile:
                presenter.OnNewBaseFileTargetLocationChanged();
                break;
        }
    }

    private void Button2_Click(object sender, EventArgs e) => presenter.SelectTargetFiles();

    private void Button3_Click(object sender, EventArgs e) => MergeClicked?.Invoke(this, EventArgs.Empty);

    private void Button4_Click(object sender, EventArgs e) => SortClicked?.Invoke(this, EventArgs.Empty);

    private void ButtonReset_Click(object sender, EventArgs e) => ResetClicked?.Invoke(this, EventArgs.Empty);

    private void ButtonOpenFile_Click(object sender, EventArgs e) => OpenFileClicked?.Invoke(this, EventArgs.Empty);

    private void ButtonOpenFolder_Click(object sender, EventArgs e) => OpenFolderClicked?.Invoke(this, EventArgs.Empty);

    private void RbUseExistingFile_CheckedChanged(object sender, EventArgs e)
    {
        if (rbUseExistingFile.Checked)
        {
            InputFileModeChanged?.Invoke(this, InputFileMode.ExistingFile);
        }
    }

    private void RbCreateNewFile_CheckedChanged(object sender, EventArgs e)
    {
        if (rbCreateNewFile.Checked)
        {
            InputFileModeChanged?.Invoke(this, InputFileMode.NewFile);
        }
    }

    private void TextBox1_Enter(object sender, EventArgs e)
    {
        if (CurrentBaseFileMode == InputFileMode.NewFile &&
            textBox1.Text == PlaceholderDuringCreateNewMode)
        {
            textBox1.Text = string.Empty;
            textBox1.ForeColor = SystemColors.WindowText;
            textBox1.Font = new Font(textBox1.Font, FontStyle.Regular);
        }
    }

    private void TextBox1_Leave(object sender, EventArgs e)
    {
        if (CurrentBaseFileMode == InputFileMode.NewFile)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                textBox1.Text = PlaceholderDuringCreateNewMode;
                textBox1.ForeColor = SystemColors.GrayText;
                textBox1.Font = new Font(textBox1.Font, FontStyle.Italic);
                presenter.OnFilenameSet(null);
            }
            else if (textBox1.Text != PlaceholderDuringCreateNewMode)
            {
                presenter.OnFilenameSet(textBox1.Text.Trim());
            }
        }
    }

    private void TextBox1_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && CurrentBaseFileMode == InputFileMode.NewFile)
        {
            e.SuppressKeyPress = true;
            this.SelectNextControl(textBox1, true, true, true, true);
        }
    }

    private async Task AnimateProgressBarAsync(ToolStripProgressBar progressBar, int steps, int delayMs, CancellationToken token)
    {
        progressBar.Value = 0;
        int maxValue = (int)(progressBar.Maximum * 0.9);
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