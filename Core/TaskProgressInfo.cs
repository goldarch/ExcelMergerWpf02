namespace ExcelMergerWpf02.Core
{
    public enum ReportLevel { Information, Warning, Error, Success, Detail, StatusUpdate, ProcessStart, ProcessEnd, ProcessCancelled }

    public class TaskProgressInfo
    {
        public string Content { get; }
        public string ProgressText { get; }
        public int? ProgressValue { get; } // 只保留这一个进度值
        public ReportLevel Level { get; }
        public object Tag { get; }

        public TaskProgressInfo(string content, string progressText, int? progressValue, ReportLevel level, object tag = null)
        {
            Content = content;
            ProgressText = progressText;
            ProgressValue = progressValue;
            Level = level;
            Tag = tag;
        }

        public TaskProgressInfo(string content, ReportLevel level, object tag = null)
            : this(content, null, null, level, tag) { }

        public TaskProgressInfo(string progressText, int? progressValue, ReportLevel level = ReportLevel.StatusUpdate, object tag = null)
            : this(null, progressText, progressValue, level, tag) { }
    }
}