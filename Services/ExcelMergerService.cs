using ExcelMergerWpf02.Core;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelDataReader;

namespace ExcelMergerWpf02.Services
{
    public class ExcelMergerService
    {
        private const int MAX_ROWS = 1048576;

        public string MergeFiles(string[] files, string destinationFile, int keyColumnIndex, int reportChunkSize, IProgress<TaskProgressInfo> progress, CancellationToken token)
        {
            // =======================================================================
            // *** 核心修改：移除仅在 .NET Core/5+ 中需要的编码注册行 ***
            // System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            // =======================================================================

            string tempFilePath = destinationFile + "." + Guid.NewGuid().ToString("N") + ".tmp";
            List<string> firstFileHeader = null;
            IWorkbook mainWorkbook = null;

            try
            {
                if (files == null || files.Length == 0) return "错误：没有提供任何需要合并的文件。";

                progress.Report(new TaskProgressInfo("任务开始...", "0%", 0, ReportLevel.ProcessStart));

                mainWorkbook = new SXSSFWorkbook();
                ISheet mainWorksheet = mainWorkbook.CreateSheet("MergedData");
                int globalRowIndex = 0;

                for (int i = 0; i < files.Length; i++)
                {
                    token.ThrowIfCancellationRequested();
                    string file = files[i];
                    progress.Report(new TaskProgressInfo($"准备处理文件: {Path.GetFileName(file)}", $"[{i + 1}/{files.Length}]", null, ReportLevel.Information, file));

                    using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            if (!reader.Read()) continue;

                            var currentHeader = new List<string>();
                            for (int j = 0; j < reader.FieldCount; j++)
                            {
                                currentHeader.Add(reader.GetValue(j)?.ToString().Trim() ?? "");
                            }

                            if (i == 0)
                            {
                                firstFileHeader = new List<string>(currentHeader);
                                IRow destHeaderRow = mainWorksheet.CreateRow(globalRowIndex++);
                                for (int j = 0; j < firstFileHeader.Count; j++)
                                {
                                    destHeaderRow.CreateCell(j).SetCellValue(firstFileHeader[j]);
                                }
                            }
                            else
                            {
                                if (!firstFileHeader.SequenceEqual(currentHeader))
                                {
                                    string errorMessage = $"标题行与首个文件不匹配。";
                                    progress.Report(new TaskProgressInfo(errorMessage, null, null, ReportLevel.Error, file));
                                    return $"文件 '{Path.GetFileName(file)}' {errorMessage}";
                                }
                            }

                            int rowsProcessedInFile = 0;
                            while (reader.Read())
                            {
                                if (globalRowIndex >= MAX_ROWS)
                                {
                                    string limitErrorMessage = $"已达到Excel最大行数限制 ({MAX_ROWS})，合并中止。";
                                    progress.Report(new TaskProgressInfo(limitErrorMessage, "Limit Reached", null, ReportLevel.Error));
                                    return limitErrorMessage;
                                }

                                token.ThrowIfCancellationRequested();
                                rowsProcessedInFile++;

                                string keyColumnValue = reader.GetValue(keyColumnIndex)?.ToString();
                                if (string.IsNullOrWhiteSpace(keyColumnValue)) continue;

                                IRow destRow = mainWorksheet.CreateRow(globalRowIndex++);
                                for (int j = 0; j < reader.FieldCount; j++)
                                {
                                    destRow.CreateCell(j).SetCellValue(reader.GetValue(j)?.ToString());
                                }

                                totalDataRowsMerged++;

                                if (rowsProcessedInFile % reportChunkSize == 0)
                                {
                                    double overallProgress = ((double)i / files.Length) + ((1.0 / files.Length) * 0.5);
                                    int overallPercentage = (int)(overallProgress * 100);
                                    string detailedContent = $"正在处理: {Path.GetFileName(file)} (已处理 {rowsProcessedInFile} 行)";
                                    string progressTextOnBar = $"[{i + 1}/{files.Length}]";
                                    progress.Report(new TaskProgressInfo(detailedContent, progressTextOnBar, overallPercentage, ReportLevel.StatusUpdate, file));
                                }
                            }
                        }
                    }
                    progress.Report(new TaskProgressInfo("已处理", null, null, ReportLevel.Detail, file));
                }

                progress.Report(new TaskProgressInfo("所有文件已读取，正在生成并保存合并文件...", "...", 99, ReportLevel.Information));
                using (var fs = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                {
                    mainWorkbook.Write(fs);
                }

                if (File.Exists(destinationFile)) File.Delete(destinationFile);
                File.Move(tempFilePath, destinationFile);

                string finalMessage = $"全部合并完成！共合并 {totalDataRowsMerged:N0} 条数据记录。";
                progress.Report(new TaskProgressInfo(finalMessage, "✓", 100, ReportLevel.Success));

                return null;
            }
            catch (OperationCanceledException)
            {
                return "任务已被用户取消。";
            }
            catch (Exception ex)
            {
                return $"发生严重错误: {ex.Message}";
            }
            finally
            {
                (mainWorkbook as IDisposable)?.Dispose();
                if (File.Exists(tempFilePath))
                {
                    try { File.Delete(tempFilePath); } catch { /* Ignore */ }
                }
            }
        }

        // I've added the totalDataRowsMerged counter back here for completeness
        private int totalDataRowsMerged = 0;
    }
}