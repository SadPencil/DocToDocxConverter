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

        private struct WorkerReport
        {
            public string Filename;
            public bool Success;
            public string Message;
        }
        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (!this.AllowDrop) return;
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
            {
                this.AllowDrop = false;
                this.MainTextBox.Visibility = Visibility.Visible;

                string[] droppedFilePaths = e.Data.GetData(DataFormats.FileDrop, true) as string[] ?? Array.Empty<string>();
                var worker = new BackgroundWorker()
                {
                    WorkerReportsProgress = true,
                };
                worker.DoWork += (object worker_sender, DoWorkEventArgs worker_e) =>
                {
                    using (var convert = new Convert())
                    {
                        for (int i = 0; i < droppedFilePaths.Length; i++)
                        {
                            var file = droppedFilePaths[i];
                            try
                            {
                                convert.ConvertFile(file);
                                (worker_sender as BackgroundWorker).ReportProgress(100 * (i + 1) / droppedFilePaths.Length, new WorkerReport
                                {
                                    Filename = file,
                                    Message = string.Empty,
                                    Success = true,
                                });
                                if (this.DeleteOriginalFileToTrash ?? false)
                                {
                                    RecycleBin.DeleteFile(file);
                                }
                            }
                            catch (Exception ex)
                            {
                                (worker_sender as BackgroundWorker).ReportProgress(100 * (i + 1) / droppedFilePaths.Length, new WorkerReport
                                {
                                    Filename = file,
                                    Message = ex.Message,
                                    Success = false,
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
                };
                worker.ProgressChanged += (object worker_sender, ProgressChangedEventArgs worker_e) =>
                {
                    MainTextAppendInfo($"{worker_e.ProgressPercentage}%");
                    var report = (WorkerReport)worker_e.UserState;
                    if (report.Success)
                    {
                        MainTextAppendInfo($"File {report.Filename} was successfully converted.");
                    }
                    else
                    {
                        MainTextAppendError($"Failed to convert file {report.Filename}. {report.Message}");
                    }
                };
                MainTextAppendInfo($"Preparing to convert {droppedFilePaths.Length} files. Initializing. Starting Word, Excel, and Powerpoint...");
                worker.RunWorkerAsync();
            }
        }

    }
}
