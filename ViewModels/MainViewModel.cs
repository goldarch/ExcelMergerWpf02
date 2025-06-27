using ExcelMergerWpf02.Commands;
using ExcelMergerWpf02.Core;
using ExcelMergerWpf02.Models;
using ExcelMergerWpf02.Services;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ExcelMergerWpf02.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        //--- Private Fields ---
        private string _outputFolder;
        private string _outputFilename = "MergedFile.xlsx";
        private int _keyColumnIndex = 0;
        private int _chunkSize = 2000;
        private readonly ExcelMergerService _mergerService;
        private readonly TaskWrapper _taskWrapper;
        private TaskProgressInfo _lastProgressReport;

        //--- Public Properties for UI Binding ---
        public ObservableCollection<FileItem> FilesToMerge { get; } = new ObservableCollection<FileItem>();
        public TaskWrapper TaskWrapper => _taskWrapper;
        public string OutputFolder { get => _outputFolder; set { _outputFolder = value; OnPropertyChanged(); } }
        public string OutputFilename { get => _outputFilename; set { _outputFilename = value; OnPropertyChanged(); } }
        public int KeyColumnIndex { get => _keyColumnIndex; set { _keyColumnIndex = value; OnPropertyChanged(); } }
        public int ChunkSize { get => _chunkSize; set { _chunkSize = value; OnPropertyChanged(); } }
        public ObservableCollection<string> LogMessages { get; } = new ObservableCollection<string>();
        public TaskProgressInfo LastProgressReport
        {
            get => _lastProgressReport;
            private set { _lastProgressReport = value; OnPropertyChanged(); }
        }

        //--- Commands for UI Binding ---
        public ICommand SelectOutputFolderCommand { get; }
        public ICommand StartMergeCommand { get; }
        public ICommand CancelMergeCommand { get; }
        public ICommand ClearListCommand { get; }

        public MainViewModel()
        {
            _mergerService = new ExcelMergerService();
            _taskWrapper = new TaskWrapper();

            _taskWrapper.ProgressReporter.ProgressChanged += OnProgressChanged;
            _taskWrapper.StateChanged += OnStateChanged;
            _taskWrapper.PropertyChanged += TaskWrapper_PropertyChanged;

            SelectOutputFolderCommand = new RelayCommand(
                p =>
                {
                    var dialog = new OpenFileDialog { CheckFileExists = false, FileName = "选择文件夹", ValidateNames = false };
                    if (dialog.ShowDialog() == true)
                    {
                        OutputFolder = Path.GetDirectoryName(dialog.FileName);
                    }
                });

            ClearListCommand = new RelayCommand(
                p => FilesToMerge.Clear(),
                p => FilesToMerge.Any() && !_taskWrapper.IsBusy
            );

            StartMergeCommand = new RelayCommand(
                async p =>
                {
                    foreach (var item in FilesToMerge) { item.Status = "等待中..."; }
                    LogMessages.Clear();
                    LastProgressReport = null;

                    var files = FilesToMerge.Select(f => f.FilePath).ToArray();
                    var finalOutputPath = Path.Combine(OutputFolder, OutputFilename);

                    _taskWrapper.DoWorkFuncAsync = (token, progress) =>
                        Task.Run(() => _mergerService.MergeFiles(files, finalOutputPath, KeyColumnIndex, ChunkSize, progress, token), token);

                    await _taskWrapper.StartTaskAsync();
                },
                p => !_taskWrapper.IsBusy && FilesToMerge.Any() && !string.IsNullOrEmpty(OutputFolder) && !string.IsNullOrEmpty(OutputFilename)
            );

            CancelMergeCommand = new RelayCommand(
                p => _taskWrapper.RequestCancel(),
                p => _taskWrapper.IsBusy
            );
        }

        private void TaskWrapper_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(TaskWrapper.IsBusy))
            {
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    (StartMergeCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (CancelMergeCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (ClearListCommand as RelayCommand)?.RaiseCanExecuteChanged();
                });
            }
        }

        private void OnProgressChanged(object sender, Core.TaskProgressInfo e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                var mergedContent = e.Content ?? _lastProgressReport?.Content;
                var mergedProgressText = e.ProgressText ?? _lastProgressReport?.ProgressText;
                var mergedProgressValue = e.ProgressValue ?? _lastProgressReport?.ProgressValue;
                var currentLevel = e.Level;
                var currentTag = e.Tag;

                LastProgressReport = new Core.TaskProgressInfo(mergedContent, mergedProgressText, mergedProgressValue, currentLevel, currentTag);

                if (e.Content != null)
                {
                    LogMessages.Insert(0, $"{DateTime.Now:HH:mm:ss} - {e.Content}");
                }

                if (e.Tag is string filePath)
                {
                    var fileItem = FilesToMerge.FirstOrDefault(f => f.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
                    if (fileItem != null)
                    {
                        switch (e.Level)
                        {
                            case ReportLevel.StatusUpdate:
                                fileItem.Status = "处理中...";
                                break;
                            case ReportLevel.Detail:
                                fileItem.Status = "已处理";
                                break;
                            case ReportLevel.Error:
                                fileItem.Status = "失败";
                                break;
                        }
                    }
                }
            });
        }

        private void OnStateChanged(object sender, Core.TaskStateChangedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                string stateMessage;
                switch (e.NewState)
                {
                    case Core.TaskExecutionState.Completed:
                        stateMessage = "[任务成功完成] 文件已保存。";
                        LastProgressReport = new Core.TaskProgressInfo("全部合并完成！", "✓", 100, Core.ReportLevel.Success);
                        foreach (var item in FilesToMerge) { item.Status = "完成"; }
                        break;

                    case Core.TaskExecutionState.Faulted:
                        stateMessage = $"[任务失败] {e.Exception?.Message}";
                        LastProgressReport = new Core.TaskProgressInfo(
                            $"错误: {e.Exception?.Message}",
                            "Error",
                            LastProgressReport?.ProgressValue,
                            Core.ReportLevel.Error
                        );
                        foreach (var item in FilesToMerge.Where(f => f.Status.Contains("处理中"))) { item.Status = "失败"; }
                        break;

                    case Core.TaskExecutionState.Cancelled:
                        stateMessage = "[任务已取消]";
                        LastProgressReport = new Core.TaskProgressInfo(
                           "任务已被用户取消",
                           "Cancelled",
                           LastProgressReport?.ProgressValue,
                           Core.ReportLevel.ProcessCancelled
                        );
                        foreach (var item in FilesToMerge.Where(f => f.Status.Contains("处理中"))) { item.Status = "已取消"; }
                        break;

                    default:
                        stateMessage = $"任务状态: {e.NewState}";
                        break;
                }
                LogMessages.Insert(0, $"{DateTime.Now:HH:mm:ss} - {stateMessage}");
            });
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}