using ExcelMergerWpf02.Models;
using ExcelMergerWpf02.ViewModels;
using System.Collections.Specialized;
using System.Linq;
using System.Windows;

namespace ExcelMergerWpf02
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            if (this.DataContext is MainViewModel viewModel)
            {
                viewModel.LogMessages.CollectionChanged += LogMessages_CollectionChanged;
            }
        }

        private void LogMessages_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems.Count > 0)
            {
                // UIElement.ScrollIntoView a thread safe call to the UI thread.
                LogListBox.ScrollIntoView(e.NewItems[0]);
            }
        }

        private void ListView_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void ListView_Drop(object sender, DragEventArgs e)
        {
            var viewModel = DataContext as MainViewModel;
            if (viewModel == null) return;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var filePaths = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (filePaths == null || filePaths.Length == 0) return;

                // =======================================================================
                // *** 新增逻辑：如果列表当前为空，则设置默认输出文件夹 ***
                // =======================================================================
                if (viewModel.FilesToMerge.Count == 0)
                {
                    // 将第一个拖入文件的目录设置为默认输出文件夹
                    viewModel.OutputFolder = System.IO.Path.GetDirectoryName(filePaths[0]);
                }
                // =======================================================================

                // 将文件添加到列表的逻辑保持不变
                foreach (var path in filePaths)
                {
                    var extension = System.IO.Path.GetExtension(path).ToLower();
                    if ((extension == ".xls" || extension == ".xlsx") && !viewModel.FilesToMerge.Any(f => f.FilePath.Equals(path, System.StringComparison.OrdinalIgnoreCase)))
                    {
                        viewModel.FilesToMerge.Add(new FileItem(path));
                    }
                }
            }
        }
    }
}