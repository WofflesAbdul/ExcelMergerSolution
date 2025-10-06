using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

public partial class FileSelectionForm : Form, IFileSelectionView
{
    private const string PlaceholderText = "Enter filename here";
    private readonly IFileSelectionPresenter presenter;
    private readonly List<Control> controlsToDisable = new List<Control>();
    private readonly List<ToolStripDropDownButton> toolStripButtonsToDisable = new List<ToolStripDropDownButton>();
    private CancellationTokenSource progressAnimationCts;

    public FileSelectionForm()
    {
        InitializeComponent(); ;
        presenter = new FileSelectionPresenter(new FileSelectionModel(), this);

        controlsToDisable.AddRange(new Control[] { button1, button2, });
        toolStripButtonsToDisable.AddRange(new ToolStripDropDownButton[] { toolStripButton1, toolStripButton2 });

        rbUseExistingFile.Checked = true;
    }

    public event EventHandler ResetRequested;

    public event EventHandler OpenFileClicked;

    public event EventHandler OpenFolderClicked;

    public event EventHandler<InputFileMode> InputFileModeChanged;

    public void DisplayFileName(string fileName)
    {
        textBox1.Text = fileName;
    }

    public void DisplayDirectoryPath(string directoryPath)
    {
        label4.Text = directoryPath;
    }

    public void DisplayTargetFilePaths(IEnumerable<string> targetFileNames)
    {
        textBox2.Text = string.Join(", ", targetFileNames);
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
    }

    public void SetOpenFolderButtonEnabled(bool enabled)
    {
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
        if (this.InvokeRequired) // 'this' = your Form
        {
            this.Invoke(new Action(() => SetProgress(percent)));
            return;
        }

        // Clamp the value between min/max
        percent = Math.Max(toolStripProgressBar1.Minimum, Math.Min(toolStripProgressBar1.Maximum, percent));
        toolStripProgressBar1.Value = percent;
    }

    public void SetOngoingStatus(string message)
    {
        toolStripLabel4.Text = message;
    }

    public void SetCompletionStatus(string message, bool isError = false)
    {
        toolStripLabel5.Text = message;
        toolStripLabel5.ForeColor = isError ? Color.Red : Color.Green;

        // Create a self-contained timer
        var timer = new System.Windows.Forms.Timer();
        timer.Interval = 5000; // 5 seconds
        timer.Tick += (s, e) =>
        {
            toolStripLabel5.Text = string.Empty;
            toolStripLabel5.ForeColor = Color.Black;
            timer.Stop();
            timer.Dispose(); // Dispose itself
        };
        timer.Start();
    }

    public async Task AnimateProgressBarAsync(int steps, int delayMs, CancellationToken token)
    {
        toolStripProgressBar1.Value = 0;
        int maxValue = (int)(toolStripProgressBar1.Maximum * 0.9);
        int stepValue = maxValue / steps;

        for (int i = 1; i <= steps; i++)
        {
            if (token.IsCancellationRequested) break;
            await Task.Delay(delayMs, token);
            if (token.IsCancellationRequested) break;
            toolStripProgressBar1.Value = Math.Min(stepValue * i, maxValue);
        }
    }

    public void NewFileCreated()
    {
        rbUseExistingFile.Checked = true;
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        toolStripLabel2.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        toolStripLabel2.AutoSize = true;
    }

    private void Button1_Click(object sender, EventArgs e)
    {
        switch (presenter.InputFileMode)
        {
            case InputFileMode.ExistingFile:
                presenter.SelectBaseFile();
                break;
            case InputFileMode.NewFile:
                presenter.SelectDirectoryPath();
                break;
        }
    }

    private void Button2_Click(object sender, EventArgs e)
    {
        presenter.SelectTargetFiles();
    }

    private async void Button3_Click(object sender, EventArgs e)
    {
        if (presenter.InputFileMode == InputFileMode.NewFile)
        {
            await presenter.RunOperationAsync(OperationRequested.CreateNewFile, presenter.CreateNewFileAction);
        }

        await presenter.RunOperationAsync(OperationRequested.Merge, presenter.MergeAction);
    }

    private async void Button4_Click(object sender, EventArgs e)
    {
        await presenter.RunOperationAsync(OperationRequested.Sort, presenter.SortAction);
    }

    private void ButtonReset_Click(object sender, EventArgs e)
    {
        ResetRequested?.Invoke(this, EventArgs.Empty);
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
            InputFileModeChanged?.Invoke(sender, InputFileMode.ExistingFile);

            textBox1.ReadOnly = true;
            textBox1.ForeColor = SystemColors.WindowText;
            textBox1.Font = new Font(textBox1.Font, FontStyle.Regular);
            button1.Text = "Select File";
        }
    }

    private void RbCreateNewFile_CheckedChanged(object sender, EventArgs e)
    {
        if (rbCreateNewFile.Checked)
        {
            InputFileModeChanged?.Invoke(sender, InputFileMode.NewFile);

            textBox1.ReadOnly = false;
            TextBox1_Leave(textBox1, EventArgs.Empty);
            button1.Text = "Select Folder";
        }
    }

    private void TextBox1_Enter(object sender, EventArgs e)
    {
        if (presenter.InputFileMode == InputFileMode.NewFile)
        {
            textBox1.Text = string.Empty;
            textBox1.ForeColor = SystemColors.WindowText;
            textBox1.Font = new Font(textBox1.Font, FontStyle.Regular);
        }
    }

    private void TextBox1_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && presenter.InputFileMode == InputFileMode.NewFile)
        {
            e.SuppressKeyPress = true;
            this.SelectNextControl(textBox1, true, true, true, true);
        }
    }

    private void TextBox1_Leave(object sender, EventArgs e)
    {
        if (presenter.InputFileMode == InputFileMode.NewFile)
        {
            presenter.OnFilenameSet(textBox1.Text.Trim());

            if (string.IsNullOrEmpty(textBox1.Text))
            {
                textBox1.Text = PlaceholderText;
                textBox1.ForeColor = SystemColors.GrayText;
                textBox1.Font = new Font(textBox1.Font, FontStyle.Italic);
            }
        }
    }
}