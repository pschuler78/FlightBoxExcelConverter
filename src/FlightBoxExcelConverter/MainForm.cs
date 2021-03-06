﻿using System;
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
        delegate void BooleanArgReturningVoidDelegate(bool value);
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
                SetButtonEnabled(false);
                if (File.Exists(textBoxExportFolderName.Text))
                {
                    DialogResult result = MessageBox.Show("Export-Datei existiert bereits. Soll die Datei überschrieben werden?",
                        "Datei überschreiben?", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk,
                        MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.No)
                    {
                        SetButtonEnabled(true);
                        return;
                    }
                }

                textBoxLog.Clear();

                _flightBoxExcelConverter = new FlightBoxExcelConverter(textBoxImportFileName.Text, textBoxExportFolderName.Text, checkBoxIgnoreDateRange.Checked);

                try
                {
                    var lastWriteDateTime = _flightBoxExcelConverter.GetLastWriteDateTimeOfImportFileName();

                    if (lastWriteDateTime.AddDays(20) < DateTime.Now)
                    {
                        var result = MessageBox.Show($"Achtung: Die zu importierende Datei {_flightBoxExcelConverter.ImportFileName} ist wahrscheinlich veraltet (Datei wurde am {lastWriteDateTime.ToShortDateString()} erstellt)!{Environment.NewLine}Soll diese Datei trotzdem verarbeitet werden?", "Warnung", MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.No)
                        {
                            SetButtonEnabled(true);
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fehler beim Prüfen der Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    SetButtonEnabled(true);
                    return;
                }

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
                SetButtonEnabled(true);
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

        private void SetButtonEnabled(bool enabled)
        {
            // InvokeRequired required compares the thread ID of the  
            // calling thread to the thread ID of the creating thread.  
            // If these threads are different, it returns true.  
            if (buttonConvert.InvokeRequired)
            {
                BooleanArgReturningVoidDelegate d = new BooleanArgReturningVoidDelegate(SetButtonEnabled);
                Invoke(d, new object[] { enabled });
            }
            else
            {
                buttonConvert.Enabled = enabled;
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
            }
            else
            {
                MessageBox.Show("Daten erfolgreich konvertiert.", "Konvertierung fertiggestellt",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            SetButtonEnabled(true);
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
