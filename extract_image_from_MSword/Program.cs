using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Microsoft.Extensions.Configuration;

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

            // 設定読み込み
            var settings = GetSettings();

            // メイン処理
            DeleteFiles(settings.ImageDir);
            var docxFiles = GetWordFiles(settings.TargetDir);
            var copyFiles = CopyWordFiles(settings.ImageDir, docxFiles);
            var extractDirs = Extract(copyFiles);
            SplitDirectory(extractDirs);
            RenameFileName(settings.ImageDir);

            sw.Stop();
            Console.WriteLine(sw.Elapsed);
        }

        /// <summary>
        /// AppSettings.jsonから値を取得します。
        /// </summary>
        /// <returns>設定情報</returns>
        private static Setting GetSettings()
        {
            var builder = new ConfigurationBuilder();
            builder.SetBasePath(Directory.GetCurrentDirectory());
            builder.AddJsonFile("AppSettings.json");
            var config = builder.Build();

            return new Setting
            {
                TargetDir = config.GetSection("AppSettings")["TargetDir"],
                ImageDir = config.GetSection("AppSettings")["ImageDir"]
            };
        }

        /// <summary>
        /// 対象ディレクトリ配下のディレクトリ・ファイルを削除します。
        /// </summary>
        private static void DeleteFiles(string targetDir)
        {
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
                return;
            }

            var target = new DirectoryInfo(targetDir);
            foreach (var file in target.GetFiles())
            {
                file.Delete();
            }

            foreach (var dir in target.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        /// <summary>
        /// 対象のディレクトリからWordファイルのリストを取得します。
        /// </summary>
        /// <param name="targetDir"></param>
        /// <returns></returns>
        private static IEnumerable<string> GetWordFiles(string targetDir)
        {
            return Directory.GetFiles(targetDir, "*.docx", SearchOption.AllDirectories);
        }

        /// <summary>
        /// docxFilesのファイルをimageDirにコピーします。
        /// </summary>
        /// <param name="imageDir"></param>
        /// <param name="docxFiles"></param>
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

                File.Copy(docxFile, destFileName, true);
            }

            return rtnList;
        }

        /// <summary>
        /// 対象のzipファイルから画像ファイルを抜き出します。
        /// </summary>
        /// <param name="zipFiles">zipファイルリスト</param>
        /// <returns>画像を格納しているパスのリスト</returns>
        private static IEnumerable<string> Extract(IEnumerable<string> zipFiles)
        {
            var rtnList = new List<string>();
            foreach (var file in zipFiles)
            {
                // docxをzipに変換
                var zipFileName = file.Replace(".docx", ".zip");
                File.Move(file, zipFileName);

                // 画像のコピー先
                var imageFilePath = file.Replace(".docx", "");
                if (!Directory.Exists(imageFilePath))
                {
                    Directory.CreateDirectory(imageFilePath);
                }

                rtnList.Add(imageFilePath);

                // 画像ファイルだけ取り出す
                using (var archive = ZipFile.OpenRead(zipFileName))
                {
                    var imageArchiveEntries = archive.Entries
                        .Where(x => x.FullName.Contains(@"word/media/image"));

                    foreach (var entry in imageArchiveEntries)
                    {
                        entry.ExtractToFile(Path.Combine(imageFilePath, entry.Name.Replace("image", "")));
                    }
                }

                // zipファイルを削除
                File.Delete(zipFileName);
            }

            return rtnList;
        }

        /// <summary>
        /// 画像ファイルを50個単位でサブディレクトリに分割します。
        /// 50個以下の場合は、サブディレクトリを作成しません。
        /// </summary>
        /// <param name="extractDirs"></param>
        private static void SplitDirectory(IEnumerable<string> extractDirs)
        {
            foreach (var dir in extractDirs)
            {
                var fileList = GetFileList(dir).ToList();
                var maxLength = fileList.Max(x => x.Length);

                // ファイルの桁数をあわせる（ファイルをソートするため）
                foreach (var file in fileList)
                {
                    if (file.Length == maxLength)
                    {
                        continue;
                    }

                    // 追加する0の数
                    var zero = string.Empty;
                    for (var i = 0; i < maxLength - file.Length; i++)
                    {
                        zero += "0";
                    }

                    var sourceFileName = Path.Combine(dir, file);
                    var destFileName = Path.Combine(dir, zero + file);
                    File.Move(sourceFileName, destFileName);
                }

                // ファイル名を変更したので、一覧を取得し直す
                fileList = GetFileList(dir).ToList();

                // 50ファイルずつ分割
                if (fileList.Count <= 50)
                {
                    continue;
                }

                foreach (var file in fileList)
                {
                    // ファイル番号
                    var num = int.Parse(Path.GetFileNameWithoutExtension(file));

                    // 格納先ディレクトリ
                    var location = Path.Combine(dir, $"{num / 50 + 1:000}");
                    if (!Directory.Exists(location))
                    {
                        Directory.CreateDirectory(location);
                    }

                    var sourceFileName = Path.Combine(dir, file);
                    var destFileName = Path.Combine(dir, location, file);
                    File.Move(sourceFileName, destFileName);
                }
            }
        }

        /// <summary>
        /// ファイル名を変更します。
        /// ・ディレクトリ単位で01から連番（51→1）
        /// ・3桁のファイル名を2桁に変更
        /// </summary>
        /// <param name="imageDir"></param>
        private static void RenameFileName(string imageDir)
        {
            var files = Directory.GetFiles(imageDir, "*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                // ファイル番号
                var num = int.Parse(Path.GetFileNameWithoutExtension(file));
                if (num >= 50)
                {
                    num++;
                }

                var dir = Path.GetDirectoryName(file);
                var fileName = $"{num % 50:00}" + Path.GetExtension(file);
                var destFileName = Path.Combine(dir, fileName);
                File.Move(file, destFileName);
            }
        }

        /// <summary>
        /// 対象のディレクトリ配下のファイル一覧を返します。
        /// ファイル名のみの一覧を返します。
        /// </summary>
        /// <param name="dir"></param>
        /// <returns></returns>
        private static IEnumerable<string> GetFileList(string dir)
        {
            return Directory.GetFiles(dir)
                .Select(Path.GetFileName)
                .OrderBy(x => x);
        }
    }

    /// <summary>
    /// 設定を保持するクラス
    /// </summary>
    internal class Setting
    {
        /// <summary>
        /// 抽出対象のwordファイルを格納しているディレクトリ
        /// </summary>
        public string TargetDir { get; set; }

        /// <summary>
        /// wordから抽出した画像ファイルを格納するディレクトリ
        /// </summary>
        public string ImageDir { get; set; }
    }
}