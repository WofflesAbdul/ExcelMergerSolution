using System;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExcelMerge
{
    /// <summary>
    /// Displays an exception summary with optional diagnostic details.
    /// </summary>
    public sealed class ExceptionDialog : Form
    {
        private readonly Exception exception;
        private bool detailsExpanded;
        private int collapsedHeight;

        private IContainer components;
        private TableLayoutPanel mainLayout;
        private Panel headingPanel;
        private Label errorIconLabel;
        private Label headingLabel;
        private TextBox messageTextBox;
        private Button detailsButton;
        private Panel detailsPanel;
        private TextBox detailsTextBox;
        private Panel buttonPanel;
        private Button copyDetailsButton;
        private Button okButton;

        /// <summary>
        /// Initializes a new instance for the Visual Studio Windows Forms Designer.
        /// </summary>
        public ExceptionDialog()
            : this(new InvalidOperationException("Example error message for Designer preview."))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExceptionDialog"/> class.
        /// </summary>
        /// <param name="exception">The exception to display.</param>
        public ExceptionDialog(Exception exception)
        {
            this.exception = exception ?? throw new ArgumentNullException(nameof(exception));
            InitializeComponent();

            messageTextBox.Text = BuildSummary(exception);
            detailsTextBox.Text = BuildDetails(exception);
            collapsedHeight = Height;
        }

        /// <summary>
        /// Shows the exception dialog and returns the selected dialog result.
        /// </summary>
        /// <param name="owner">The owner window.</param>
        /// <param name="exception">The exception to display.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(IWin32Window owner, Exception exception)
        {
            using (var dialog = new ExceptionDialog(exception))
            {
                return owner == null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
            }
        }

        private void InitializeComponent()
        {
            components = new Container();
            mainLayout = new TableLayoutPanel();
            headingPanel = new Panel();
            errorIconLabel = new Label();
            headingLabel = new Label();
            messageTextBox = new TextBox();
            detailsButton = new Button();
            detailsPanel = new Panel();
            detailsTextBox = new TextBox();
            buttonPanel = new Panel();
            copyDetailsButton = new Button();
            okButton = new Button();

            SuspendLayout();

            // Form
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Control;
            CancelButton = okButton;
            ClientSize = new Size(540, 190);
            Font = SystemFonts.MessageBoxFont;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ExceptionDialog";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "Error";
            KeyPreview = true;
            KeyDown += ExceptionDialog_KeyDown;

            // Main layout
            mainLayout.ColumnCount = 1;
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            mainLayout.Controls.Add(headingPanel, 0, 0);
            mainLayout.Controls.Add(messageTextBox, 0, 1);
            mainLayout.Controls.Add(detailsButton, 0, 2);
            mainLayout.Controls.Add(detailsPanel, 0, 3);
            mainLayout.Controls.Add(buttonPanel, 0, 4);
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.Location = new Point(0, 0);
            mainLayout.Name = "mainLayout";
            mainLayout.Padding = new Padding(12);
            mainLayout.RowCount = 5;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 0F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            mainLayout.TabIndex = 0;

            // Heading
            headingPanel.Controls.Add(errorIconLabel);
            headingPanel.Controls.Add(headingLabel);
            headingPanel.Dock = DockStyle.Fill;
            headingPanel.Location = new Point(15, 15);
            headingPanel.Margin = new Padding(0, 0, 0, 4);
            headingPanel.Name = "headingPanel";
            headingPanel.TabIndex = 0;

            errorIconLabel.BorderStyle = BorderStyle.FixedSingle;
            errorIconLabel.Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 17F, FontStyle.Bold);
            errorIconLabel.ForeColor = Color.Firebrick;
            errorIconLabel.Location = new Point(0, 2);
            errorIconLabel.Name = "errorIconLabel";
            errorIconLabel.Size = new Size(28, 28);
            errorIconLabel.TabIndex = 0;
            errorIconLabel.Text = "!";
            errorIconLabel.TextAlign = ContentAlignment.MiddleCenter;

            headingLabel.AutoSize = true;
            headingLabel.Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold);
            headingLabel.Location = new Point(40, 7);
            headingLabel.Name = "headingLabel";
            headingLabel.TabIndex = 1;
            headingLabel.Text = "The operation could not be completed.";

            // Error summary
            messageTextBox.BackColor = SystemColors.Control;
            messageTextBox.BorderStyle = BorderStyle.None;
            messageTextBox.Dock = DockStyle.Fill;
            messageTextBox.Location = new Point(15, 57);
            messageTextBox.Margin = new Padding(0);
            messageTextBox.Multiline = true;
            messageTextBox.Name = "messageTextBox";
            messageTextBox.ReadOnly = true;
            messageTextBox.ScrollBars = ScrollBars.Vertical;
            messageTextBox.TabIndex = 1;
            messageTextBox.Text = "Exception type and message appear here.";

            // Expand/collapse details
            detailsButton.Anchor = AnchorStyles.Left;
            detailsButton.AutoSize = true;
            detailsButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            detailsButton.Location = new Point(15, 100);
            detailsButton.Margin = new Padding(0, 4, 0, 4);
            detailsButton.Name = "detailsButton";
            detailsButton.TabIndex = 2;
            detailsButton.Text = "⌄ Show details";
            detailsButton.UseVisualStyleBackColor = true;
            detailsButton.Click += DetailsButton_Click;

            // Diagnostic details
            detailsPanel.Controls.Add(detailsTextBox);
            detailsPanel.Dock = DockStyle.Fill;
            detailsPanel.Location = new Point(15, 136);
            detailsPanel.Margin = new Padding(0, 4, 0, 4);
            detailsPanel.Name = "detailsPanel";
            detailsPanel.TabIndex = 3;
            detailsPanel.Visible = false;

            detailsTextBox.Dock = DockStyle.Fill;
            detailsTextBox.Font = new Font(FontFamily.GenericMonospace, 9F);
            detailsTextBox.Location = new Point(0, 0);
            detailsTextBox.Multiline = true;
            detailsTextBox.Name = "detailsTextBox";
            detailsTextBox.ReadOnly = true;
            detailsTextBox.ScrollBars = ScrollBars.Both;
            detailsTextBox.WordWrap = false;
            detailsTextBox.TabIndex = 0;

            // Bottom buttons
            buttonPanel.Controls.Add(copyDetailsButton);
            buttonPanel.Controls.Add(okButton);
            buttonPanel.Dock = DockStyle.Fill;
            buttonPanel.Location = new Point(15, 136);
            buttonPanel.Margin = new Padding(0, 4, 0, 0);
            buttonPanel.Name = "buttonPanel";
            buttonPanel.TabIndex = 4;
            buttonPanel.Resize += ButtonPanel_Resize;

            copyDetailsButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            copyDetailsButton.Location = new Point(328, 2);
            copyDetailsButton.Name = "copyDetailsButton";
            copyDetailsButton.Size = new Size(100, 28);
            copyDetailsButton.TabIndex = 0;
            copyDetailsButton.Text = "Copy details";
            copyDetailsButton.UseVisualStyleBackColor = true;
            copyDetailsButton.Visible = false;
            copyDetailsButton.Click += CopyDetailsButton_Click;

            okButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            okButton.DialogResult = DialogResult.OK;
            okButton.Location = new Point(436, 2);
            okButton.Name = "okButton";
            okButton.Size = new Size(80, 28);
            okButton.TabIndex = 1;
            okButton.Text = "OK";
            okButton.UseVisualStyleBackColor = true;
            okButton.Click += OkButton_Click;

            AcceptButton = okButton;
            Controls.Add(mainLayout);
            ResumeLayout(false);
        }

        private static string BuildSummary(Exception ex)
        {
            return ex.GetType().Name + Environment.NewLine + ex.Message;
        }

        private static string BuildDetails(Exception ex)
        {
            var builder = new StringBuilder();
            var current = ex;
            var level = 0;

            while (current != null)
            {
                if (level > 0)
                {
                    builder.AppendLine();
                    builder.AppendLine(new string('-', 72));
                    builder.AppendLine();
                }

                builder.AppendLine("Level: " + level);
                builder.AppendLine("Type: " + current.GetType().FullName);
                builder.AppendLine("Message: " + current.Message);
                builder.AppendLine("HRESULT: 0x" + current.HResult.ToString("X8"));
                builder.AppendLine("Stack trace:");
                builder.AppendLine(string.IsNullOrWhiteSpace(current.StackTrace) ? "(No stack trace available.)" : current.StackTrace);

                current = current.InnerException;
                level++;
            }

            return builder.ToString();
        }

        private void DetailsButton_Click(object sender, EventArgs e)
        {
            detailsExpanded = !detailsExpanded;
            detailsPanel.Visible = detailsExpanded;
            copyDetailsButton.Visible = detailsExpanded;
            detailsButton.Text = detailsExpanded ? "⌃ Hide details" : "⌄ Show details";
            mainLayout.RowStyles[3].Height = detailsExpanded ? 190F : 0F;
            Height = detailsExpanded ? collapsedHeight + 205 : collapsedHeight;
            UpdateButtonPositions();
        }

        private void CopyDetailsButton_Click(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(BuildDetails(exception));
                copyDetailsButton.Text = "Copied";
            }
            catch (System.Runtime.InteropServices.ExternalException)
            {
                MessageBox.Show(this, "Unable to copy details to the clipboard.", "Copy failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ExceptionDialog_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                try
                {
                    Clipboard.SetText(detailsExpanded ? BuildDetails(exception) : BuildSummary(exception));
                    e.Handled = true;
                }
                catch (System.Runtime.InteropServices.ExternalException)
                {
                    // Clipboard is unavailable; leave the dialog open.
                }
            }
        }

        private void ButtonPanel_Resize(object sender, EventArgs e)
        {
            UpdateButtonPositions();
        }

        private void UpdateButtonPositions()
        {
            if (buttonPanel == null || okButton == null || copyDetailsButton == null)
            {
                return;
            }

            okButton.Location = new Point(buttonPanel.ClientSize.Width - okButton.Width, 2);
            copyDetailsButton.Location = new Point(okButton.Left - copyDetailsButton.Width - 8, 2);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}