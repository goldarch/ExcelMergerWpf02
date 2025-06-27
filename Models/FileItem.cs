using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMergerWpf02.Models
{
    public class FileItem : INotifyPropertyChanged
    {
        private string _status = "等待中";

        public string FilePath { get; }
        public string FileName => Path.GetFileName(FilePath);

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        public FileItem(string filePath)
        {
            FilePath = filePath;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}