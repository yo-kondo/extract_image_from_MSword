using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;

namespace extract_image_from_MSword
{
    /// <summary>
    /// メインクラス
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// エントリポイント
        /// </summary>
        private static void Main()
        {
            var sw = new Stopwatch();
            sw.Start();

            var targetDir = ConfigurationManager.AppSettings["target_dir"];
            var imageDir = ConfigurationManager.AppSettings["image_dir"];

            var docxFiles = GetWordFiles(targetDir);
            var copyFiles = CopyWordFiles(imageDir, docxFiles);
            var zipFiles = ChangeZip(copyFiles);

            Debug.WriteLine(WriteList(zipFiles));

            sw.Stop();
            Debug.WriteLine(sw.Elapsed);
        }

        /// <summary>
        /// 対象のディレクトリからWordファイルのリスト取得します。
        /// </summary>
        private static IEnumerable<string> GetWordFiles(string targetDir)
        {
            return Directory.GetFiles(targetDir, "*.docx", SearchOption.AllDirectories);
        }

        /// <summary>
        /// docxFilesのファイルをimageDirにコピーします。
        /// </summary>
        /// <returns>コピー先のパス</returns>
        private static IEnumerable<string> CopyWordFiles(string imageDir, IEnumerable<string> docxFiles)
        {
            var rtnList = new List<string>();

            foreach (var docxFile in docxFiles)
            {
                // ファイル名
                var fileName = Path.GetFileName(docxFile);
                var destFileName = Path.Combine(imageDir, fileName);
                rtnList.Add(destFileName);

                // string sourceFileName, string destFileName)
                File.Copy(docxFile, destFileName, true);
            }
            return rtnList;
        }

        /// <summary>
        /// 対象のファイルをzipに変換します。
        /// 拡張子を変更するだけです。
        /// </summary>
        private static IEnumerable<string> ChangeZip(IEnumerable<string> targetFiles)
        {
            var rtnList = new List<string>();
            foreach (var targetFile in targetFiles)
            {
                var destFileName = targetFile.Replace(".docx", ".zip");
                rtnList.Add(destFileName);
                File.Move(targetFile, destFileName);
            }
            return rtnList;
        }

        /// <summary>
        /// Listを整形した文字列に変換して返します。
        /// </summary>
        private static string WriteList(object list)
        {
            return JsonConvert.SerializeObject(list, Formatting.Indented);
        }
    }
}
