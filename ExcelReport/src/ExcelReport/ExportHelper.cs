using ExcelReport.Renderers;
using System;
using System.IO;

namespace ExcelReport
{
    public static class ExportHelper
    {
        public static void ExportToLocal(string templateFile, string targetFile, params SheetRenderer[] sheetRenderers)
        {
            if (string.IsNullOrWhiteSpace(templateFile))
            {
                throw new ArgumentNullException("templateFile");
            }
            if (string.IsNullOrWhiteSpace(targetFile))
            {
                throw new ArgumentNullException("targetFile");
            }
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException(templateFile + " 文件不存在!");
            }

            using (FileStream fs = File.OpenWrite(targetFile))
            {
                byte[] buffer = Export.ExportToBuffer(templateFile, sheetRenderers);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
            }
        }
        /// <summary>
        /// 对ExportToLocal的增强(避免使用Foreach产生数据输出不全的情况）
        /// </summary>
        /// <param name="templateFile"></param>
        /// <param name="targetFile"></param>
        /// <param name="sheetRenderers"></param>
        public static void TExportToLocal(string templateFile, string targetFile, params SheetRenderer[] sheetRenderers)
        {
            if (string.IsNullOrWhiteSpace(templateFile))
            {
                throw new ArgumentNullException("templateFile");
            }
            if (string.IsNullOrWhiteSpace(targetFile))
            {
                throw new ArgumentNullException("targetFile");
            }
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException(templateFile + " 文件不存在!");
            }

            using (FileStream fs = File.OpenWrite(targetFile))
            {
                byte[] buffer = Export.TExportToBuffer(templateFile, sheetRenderers);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
            }
        }

    }
}