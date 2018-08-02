using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using FlightBoxExcelConverter.Properties;

namespace FlightBoxExcelConverter
{
    public partial class MainForm : Form
    {
        delegate void StringArgReturningVoidDelegate(string text);
        private FlightBoxExcelConverter _flightBoxExcelConverter;

        public MainForm()
        {
            InitializeComponent();

            if (Settings.Default.DefaultImportFileName.ToLower().EndsWith(".csv"))
            {
                if (File.Exists(Settings.Default.DefaultImportFileName))
                    textBoxImportFileName.Text = Settings.Default.DefaultImportFileName;
            }
            else
            {
                textBoxImportFileName.Text = Settings.Default.DefaultImportFileName;
            }

            textBoxExportFolderName.Text = Settings.Default.DefaultExportFolderName;
        }
        
        private void buttonBrowseImportFile_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Comma separated file|*.csv";
            openFileDialog.Title = "Import file";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxImportFileName.Text = openFileDialog.FileName;
            }
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            try
            {
                buttonConvert.Enabled = false;
                if (File.Exists(textBoxExportFolderName.Text))
                {
                    DialogResult result = MessageBox.Show("Export-Datei existiert bereits. Soll die Datei überschrieben werden?",
                        "Datei überschreiben?", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk,
                        MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.No)
                    {
                        return;
                    }
                }

                textBoxLog.Clear();

                _flightBoxExcelConverter = new FlightBoxExcelConverter(textBoxImportFileName.Text, textBoxExportFolderName.Text);
                _flightBoxExcelConverter.ExportFinished += OnExportFinished;
                _flightBoxExcelConverter.LogEventRaised += OnLogEventRaised;
                Thread t = new Thread(new ThreadStart(RunConverter));
                // start the thread using the t-variable:
                t.Start();
            }
            catch (Exception exception)
            {
                MessageBox.Show($"Fehler beim Konvertieren: {exception.Message}", "Fehler", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                buttonConvert.Enabled = true;
            }
        }

        private void RunConverter()
        {
            _flightBoxExcelConverter.Convert();
        }

        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the  
            // calling thread to the thread ID of the creating thread.  
            // If these threads are different, it returns true.  
            if (textBoxLog.InvokeRequired)
            {
                StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(SetText);
                Invoke(d, new object[] { text });
            }
            else
            {
                textBoxLog.Text += text;
            }
        }

        private void OnLogEventRaised(object sender, LogEventArgs logEventArgs)
        {
            SetText($"{logEventArgs.Text}{Environment.NewLine}");
        }

        private void OnExportFinished(object sender, EventArgs eventArgs)
        {
            _flightBoxExcelConverter.ExportFinished -= OnExportFinished;
            _flightBoxExcelConverter.LogEventRaised -= OnLogEventRaised;

            if (_flightBoxExcelConverter.HasExportError)
            {
                MessageBox.Show($"Fehler beim Konvertieren der Daten.{Environment.NewLine}{Environment.NewLine}Fehler-Meldung:{Environment.NewLine}{_flightBoxExcelConverter.ExportErrorMessage}", "Fehler",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("Daten erfolgreich konvertiert.", "Konvertierung fertiggestellt",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            buttonConvert.Enabled = true;
        }

        private void buttonBrowseExportFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBoxExportFolderName.Text = dialog.SelectedPath;
            }
        }
    }
}
