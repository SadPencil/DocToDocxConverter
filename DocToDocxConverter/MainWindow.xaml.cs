using System;
using Path = System.IO.Path;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;

namespace DocToDocxConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        public bool? DeleteOriginalFileToTrash { get; set; } = false;
        public bool? HideOfficeAppWindow { get; set; } = false;

        private void MainTextAppendError(string text) => this.MainTextAppend(text, Brushes.Red, FontWeights.Bold);
        private void MainTextAppendSpecial(string text) => this.MainTextAppend(text, Brushes.Green, FontWeights.Regular);
        private void MainTextAppendInfo(string text) => this.MainTextAppend(text, Brushes.Black, FontWeights.Regular);
        private void MainTextAppend(string text, Brush brush, FontWeight fontWeight)
        {
            var rangeOfText = new TextRange(this.MainTextBox.Document.ContentEnd, this.MainTextBox.Document.ContentEnd)
            {
                Text = text + Environment.NewLine
            };
            rangeOfText.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
            rangeOfText.ApplyPropertyValue(TextElement.FontWeightProperty, fontWeight);

            // set the current caret position to the end
            this.MainTextBox.CaretPosition = this.MainTextBox.Document.ContentEnd;
            // scroll it automatically
            this.MainTextBox.ScrollToEnd();
        }

        private class WorkerReport
        {
            public bool IsError = false;
            public string Message;
        }
        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (!this.AllowDrop) return;
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
            {
                this.AllowDrop = false;
                this.OptionsGroupBox.IsEnabled = false;
                this.MainTextBox.Visibility = Visibility.Visible;

                string[] droppedFilePaths = e.Data.GetData(DataFormats.FileDrop, true) as string[] ?? Array.Empty<string>();
                var worker = new BackgroundWorker()
                {
                    WorkerReportsProgress = true,
                };
                worker.DoWork += (object worker_sender, DoWorkEventArgs worker_e) =>
                {
                    using (var convert = new Convert(this.HideOfficeAppWindow ?? false))
                    {
                        for (int i = 0; i < droppedFilePaths.Length; i++)
                        {
                            var file = droppedFilePaths[i];
                            try
                            {
                                string newFile = convert.GetConvertedFilePath(file);
                                if (File.Exists(newFile))
                                {
                                    (worker_sender as BackgroundWorker).ReportProgress(100 * (i + 1) / droppedFilePaths.Length, new WorkerReport
                                    {
                                        Message = $"Existing file detected. Deleting. {file}",
                                    });
                                    RecycleBin.DeleteFile(newFile);
                                }

                                convert.ConvertFile(file);
                                (worker_sender as BackgroundWorker).ReportProgress(100 * (i + 1) / droppedFilePaths.Length, new WorkerReport
                                {
                                    Message = $"Successfully converted {file}.",
                                });
                                if (this.DeleteOriginalFileToTrash ?? false)
                                {
                                    (worker_sender as BackgroundWorker).ReportProgress(100 * (i + 1) / droppedFilePaths.Length, new WorkerReport
                                    {
                                        Message = $"Deleting original file {file}.",
                                    });
                                    RecycleBin.DeleteFile(file);
                                }
                            }
                            catch (Exception ex)
                            {
                                (worker_sender as BackgroundWorker).ReportProgress(100 * (i + 1) / droppedFilePaths.Length, new WorkerReport
                                {
                                    IsError = true,
                                    Message = $"An error occured while converting file {file}. {ex.Message}",
                                });
                            }
                        }

                    }

                };
                worker.RunWorkerCompleted += (object worker_sender, RunWorkerCompletedEventArgs worker_e) =>
                {
                    if (worker_e.Error == null)
                    {
                        MainTextAppendInfo($"Complete. Please view the messages above to check if there are any errors.");
                    }
                    else
                    {
                        MainTextAppendError(worker_e.Error.Message);
                        MessageBox.Show(this, worker_e.Error.Message, "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    this.AllowDrop = true;
                    this.OptionsGroupBox.IsEnabled = true;
                };
                worker.ProgressChanged += (object worker_sender, ProgressChangedEventArgs worker_e) =>
                {
                    var report = (WorkerReport)worker_e.UserState;
                    if (report.IsError)
                    {
                        MainTextAppendError($"[{worker_e.ProgressPercentage}%] {report.Message}");
                    }
                    else
                    {
                        MainTextAppendInfo($"[{worker_e.ProgressPercentage}%] {report.Message}");
                    }
                };
                MainTextAppendInfo($"Preparing to convert {droppedFilePaths.Length} files. Initializing. Starting Word, Excel, and PowerPoint...");
                worker.RunWorkerAsync();
            }
        }

    }
}
