using System.Diagnostics;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordToPDF;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

public sealed class MainForm : Form
{
    private readonly Label status;
    private readonly TextBox pathBox;
    private readonly Button pasteButton;
    private readonly Button chooseButton;
    private readonly Button convertButton;
    private readonly Button copyButton;
    private string? currentPdf;

    public MainForm()
    {
        Text = "Word → PDF";
        Width = 620;
        Height = 390;
        MinimumSize = new Size(620, 390);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);

        var title = new Label {
            Text = "Word → PDF",
            Font = new Font("Segoe UI", 22, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(32, 25)
        };
        Controls.Add(title);

        var help = new Label {
            Text = "Paste a Word file, choose it, or drag it onto this window.",
            AutoSize = true,
            Location = new Point(35, 72)
        };
        Controls.Add(help);

        pathBox = new TextBox {
            Location = new Point(35, 112),
            Width = 430,
            Height = 35,
            ReadOnly = true,
            AllowDrop = true
        };
        pathBox.DragEnter += (_, e) => {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        };
        pathBox.DragDrop += (_, e) => {
            var paths = (string[])e.Data!.GetData(DataFormats.FileDrop)!;
            if (paths.Length > 0) SetFile(paths[0]);
        };
        Controls.Add(pathBox);

        pasteButton = MakeButton("Paste", 480, 110, 100, 38);
        pasteButton.Click += (_, _) => PasteFile();
        Controls.Add(pasteButton);

        chooseButton = MakeButton("Choose Word file", 35, 170, 170, 45);
        chooseButton.Click += (_, _) => ChooseFile();
        Controls.Add(chooseButton);

        convertButton = MakeButton("Convert to PDF", 220, 170, 170, 45);
        convertButton.Enabled = false;
        convertButton.Click += (_, _) => ConvertToPdf();
        Controls.Add(convertButton);

        copyButton = MakeButton("Copy PDF", 405, 170, 170, 45);
        copyButton.Enabled = false;
        copyButton.Click += (_, _) => CopyPdf();
        Controls.Add(copyButton);

        status = new Label {
            Text = "Ready.",
            Location = new Point(35, 245),
            Width = 540,
            Height = 70,
            ForeColor = Color.DimGray
        };
        Controls.Add(status);

        AllowDrop = true;
        DragEnter += (_, e) => {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        };
        DragDrop += (_, e) => {
            var paths = (string[])e.Data!.GetData(DataFormats.FileDrop)!;
            if (paths.Length > 0) SetFile(paths[0]);
        };
    }

    private Button MakeButton(string text, int x, int y, int w, int h) =>
        new() { Text = text, Location = new Point(x, y), Width = w, Height = h };

    private void PasteFile()
    {
        try
        {
            if (Clipboard.ContainsFileDropList())
            {
                var list = Clipboard.GetFileDropList();
                if (list.Count > 0) SetFile(list[0]);
                return;
            }
            MessageBox.Show("In Word, copy the .doc or .docx file itself (for example from File Explorer), then press Paste.",
                "Paste a Word file", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex) { ShowError(ex); }
    }

    private void ChooseFile()
    {
        using var dlg = new OpenFileDialog {
            Filter = "Word documents|*.docx;*.doc|All files|*.*",
            Title = "Choose a Word document"
        };
        if (dlg.ShowDialog() == DialogResult.OK) SetFile(dlg.FileName);
    }

    private void SetFile(string file)
    {
        if (!File.Exists(file)) return;
        var ext = Path.GetExtension(file).ToLowerInvariant();
        if (ext != ".docx" && ext != ".doc")
        {
            MessageBox.Show("Please choose a .doc or .docx Word document.",
                "Not a Word file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        pathBox.Text = file;
        convertButton.Enabled = true;
        copyButton.Enabled = false;
        currentPdf = null;
        status.Text = "Word file selected. Press “Convert to PDF”.";
    }

    private void ConvertToPdf()
    {
        if (string.IsNullOrWhiteSpace(pathBox.Text)) return;

        string source = pathBox.Text;
        string pdf = Path.Combine(
            Path.GetDirectoryName(source)!,
            Path.GetFileNameWithoutExtension(source) + ".pdf");

        object missing = Type.Missing;
        Word.Application? app = null;
        Word.Document? doc = null;

        try
        {
            status.Text = "Converting…";
            Cursor = Cursors.WaitCursor;
            app = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
            doc = app.Documents.Open(source, ReadOnly: true, Visible: false);
            doc.ExportAsFixedFormat(pdf, Word.WdExportFormat.wdExportFormatPDF);
            currentPdf = pdf;
            copyButton.Enabled = true;
            status.Text = $"Done!\r\nPDF saved next to the Word file:\r\n{pdf}";
        }
        catch (COMException)
        {
            MessageBox.Show("Microsoft Word could not be opened. This app needs desktop Microsoft Word installed.",
                "Microsoft Word required", MessageBoxButtons.OK, MessageBoxIcon.Error);
            status.Text = "Conversion failed.";
        }
        catch (Exception ex)
        {
            ShowError(ex);
            status.Text = "Conversion failed.";
        }
        finally
        {
            if (doc != null) { doc.Close(false); Marshal.FinalReleaseComObject(doc); }
            if (app != null) { app.Quit(false); Marshal.FinalReleaseComObject(app); }
            Cursor = Cursors.Default;
        }
    }

    private void CopyPdf()
    {
        if (string.IsNullOrWhiteSpace(currentPdf) || !File.Exists(currentPdf)) return;

        // Windows Explorer/WhatsApp can receive the PDF as a file from the clipboard.
        var files = new System.Collections.Specialized.StringCollection { currentPdf };
        Clipboard.SetFileDropList(files);
        status.Text = "PDF copied. Open WhatsApp and paste (Ctrl+V) into the chat.";
    }

    private void ShowError(Exception ex) =>
        MessageBox.Show(ex.Message, "Something went wrong", MessageBoxButtons.OK, MessageBoxIcon.Error);
}
