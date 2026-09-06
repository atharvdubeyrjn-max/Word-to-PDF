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

    // Preferred folder on your dad's PC.
    // If it doesn't exist, the app falls back to the user's Documents folder.
    private static readonly string PreferredWordFolder =
        @"E:\Manoj Word Documents\Manoj A to Z";

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
            Text = "Choose a Word file, paste a copied file, or drag it onto this window.",
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

            MessageBox.Show(
                "Copy the Word document file from File Explorer, then press Paste.",
                "Paste a Word file",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            ShowError(ex);
        }
    }

    private void ChooseFile()
    {
        string initialFolder = GetInitialWordFolder();

        using var dlg = new OpenFileDialog {
            Filter = "Word documents|*.docx;*.doc|All files|*.*",
            Title = "Choose a Word document",
            InitialDirectory = initialFolder,
            CheckFileExists = true,
            Multiselect = false
        };

        if (dlg.ShowDialog(this) == DialogResult.OK)
            SetFile(dlg.FileName);
    }

    private static string GetInitialWordFolder()
    {
        if (Directory.Exists(PreferredWordFolder))
            return PreferredWordFolder;

        string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        string fallback = Path.Combine(
            documents,
            "Manoj Word Documents",
            "Manoj A to Z");

        if (Directory.Exists(fallback))
            return fallback;

        return Directory.Exists(documents) ? documents : Environment.CurrentDirectory;
    }

    private void SetFile(string file)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
            {
                MessageBox.Show(
                    "The selected Word file could not be found.",
                    "File not found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext != ".docx" && ext != ".doc")
            {
                MessageBox.Show(
                    "Please choose a .doc or .docx Word document.",
                    "Not a Word file",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            pathBox.Text = file;
            convertButton.Enabled = true;
            copyButton.Enabled = false;
            currentPdf = null;
            status.Text = "Word file selected. Press “Convert to PDF”.";
        }
        catch (Exception ex)
        {
            ShowError(ex);
        }
    }

    private void ConvertToPdf()
    {
        string source = pathBox.Text;

        if (string.IsNullOrWhiteSpace(source) || !File.Exists(source))
        {
            MessageBox.Show(
                "The selected Word file no longer exists. Please choose it again.",
                "File not found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        string? directory = Path.GetDirectoryName(source);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
        {
            MessageBox.Show(
                "The folder containing the Word file could not be found.",
                "Folder not found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        string pdf = Path.Combine(
            directory,
            Path.GetFileNameWithoutExtension(source) + ".pdf");

        Word.Application? app = null;
        Word.Document? doc = null;

        try
        {
            status.Text = "Converting…";
            Cursor = Cursors.WaitCursor;
            convertButton.Enabled = false;

            app = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };

            doc = app.Documents.Open(
                FileName: source,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);

            doc.ExportAsFixedFormat(
                OutputFileName: pdf,
                ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range: Word.WdExportRange.wdExportAllDocument,
                From: 0,
                To: 0,
                Item: Word.WdExportItem.wdExportDocumentContent,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                DocStructureTags: true,
                BitmapMissingFonts: true,
                UseISO19005_1: false);

            if (!File.Exists(pdf))
                throw new FileNotFoundException(
                    "Microsoft Word did not create the PDF file.",
                    pdf);

            currentPdf = pdf;
            copyButton.Enabled = true;
            status.Text = $"Done!\r\nPDF saved next to the Word file:\r\n{pdf}";
        }
        catch (COMException ex)
        {
            MessageBox.Show(
                "Microsoft Word could not convert the document.\r\n\r\n" +
                "Please make sure desktop Microsoft Word is installed and try again.\r\n\r\n" +
                $"Details: {ex.Message}",
                "Microsoft Word required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            status.Text = "Conversion failed.";
        }
        catch (UnauthorizedAccessException)
        {
            MessageBox.Show(
                "Windows did not allow the app to create the PDF in this folder. " +
                "Try saving the Word document in a folder where you have permission.",
                "Permission problem",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            status.Text = "Conversion failed.";
        }
        catch (IOException ex)
        {
            MessageBox.Show(
                "Windows could not create or access the PDF file.\r\n\r\n" +
                $"Details: {ex.Message}",
                "File error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            status.Text = "Conversion failed.";
        }
        catch (Exception ex)
        {
            ShowError(ex);
            status.Text = "Conversion failed.";
        }
        finally
        {
            if (doc != null)
            {
                try { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
                try { Marshal.FinalReleaseComObject(doc); }
                catch { }
            }

            if (app != null)
            {
                try { app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
                try { Marshal.FinalReleaseComObject(app); }
                catch { }
            }

            Cursor = Cursors.Default;
            convertButton.Enabled = !string.IsNullOrWhiteSpace(pathBox.Text) && File.Exists(pathBox.Text);
        }
    }

    private void CopyPdf()
    {
        if (string.IsNullOrWhiteSpace(currentPdf) || !File.Exists(currentPdf))
        {
            MessageBox.Show(
                "The PDF could not be found. Convert the Word file again.",
                "PDF not found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        try
        {
            var files = new System.Collections.Specialized.StringCollection
            {
                currentPdf
            };

            Clipboard.SetFileDropList(files);
            status.Text = "PDF copied. Open WhatsApp and paste (Ctrl+V) into the chat.";
        }
        catch (Exception ex)
        {
            ShowError(ex);
        }
    }

    private void ShowError(Exception ex) =>
        MessageBox.Show(
            ex.Message,
            "Something went wrong",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
}
