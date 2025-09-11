using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

public partial class Form1 : Form, IFileSelectionView
{
    private readonly FileSelectionPresenter presenter;
    private readonly List<Control> controlsToDisable = new List<Control>();
    private readonly List<ToolStripButton> toolStripButtonsToDisable = new List<ToolStripButton>();
    private Dictionary<Control, bool> controlsStateBackup;
    private Dictionary<ToolStripButton, bool> toolStripStateBackup;

    public event EventHandler MergeClicked;

    public event EventHandler SortClicked;

    public event EventHandler ResetClicked;

    public Form1()
    {
        InitializeComponent();
        presenter = new FileSelectionPresenter(this);

        controlsToDisable.AddRange(new Control[] { button1, button2, button3, button4, });
        toolStripButtonsToDisable.AddRange(new ToolStripButton[] { toolStripButton1, });

        // Subscribe to requested events
        MergeClicked += (s, e) => presenter.OnMergeClicked();
        SortClicked += (s, e) => presenter.OnSortClicked();
        ResetClicked += (s, e) => presenter.OnResetClicked();
        presenter.MergeRequested += async (s, e) => await RunMergeAsync(); // Task should do something with toolstripProgressBar
        presenter.SortRequested += async (s, e) => await RunSortAsync();// Task should do something with toolstripProgressBar
        presenter.ResetRequested += (s, e) => presenter.ResetSelection();
    }

    public void UpdateBaseFileName(string name)
    {
        textBox1.Text = name;
        presenter.CheckMergeButtonState();
        presenter.CheckSortButtonState();
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

            toolStripStateBackup = new Dictionary<ToolStripButton, bool>();
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

    private void Form1_Load(object sender, EventArgs e)
    {
        toolStripLabel2.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        toolStripLabel2.AutoSize = true;
    }

    private void Button1_Click(object sender, EventArgs e)
    {
        presenter.SelectBaseFile();
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

    private void ToolStripButton1_Click(object sender, EventArgs e)
    {
        ResetClicked?.Invoke(this, EventArgs.Empty);
    }

    private async Task RunMergeAsync()
    {
        ApplyTaskControlLock(false);
        presenter.SetProgress(0); // reset

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

            presenter.NotifyMergeCompleted();
            MessageBox.Show("Merge completed!");
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
            _ = AnimateProgressBarAsync(toolStripProgressBar1, 20, 500);

            var sorter = new Excel_Handling.ExcelSorter();
            await Task.Run(() => sorter.SortSheets(presenter.BaseFilePath));

            toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
            MessageBox.Show("Sort completed!");
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

    private async Task AnimateProgressBarAsync(ToolStripProgressBar progressBar, int steps, int delayMs)
    {
        progressBar.Value = 0;
        int maxValue = (int)(progressBar.Maximum * 0.9); // cap at 90%
        int stepValue = maxValue / steps;

        for (int i = 1; i <= steps; i++)
        {
            await Task.Delay(delayMs);
            progressBar.Value = Math.Min(stepValue * i, maxValue);
        }
    }
}
