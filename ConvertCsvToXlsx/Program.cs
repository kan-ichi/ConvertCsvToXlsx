using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace ConvertCsvToXlsx
{
    class Program
    {
        #region エントリーポイント

        /// <summary>
        /// エントリーポイント
        /// </summary>
        public static void Main(string[] _args)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;

            Request request = new Request();

            request.InputCsvFolderName = ConfigurationManager.AppSettings["InputCsvFolderName"];
            request.InputCsvEncoding = ConfigurationManager.AppSettings["InputCsvEncoding"];
            request.InputCsvSkipFirstRow = (ConfigurationManager.AppSettings["InputCsvSkipFirstRow"] == "Yes");
            request.OutputXlsxFolderName = ConfigurationManager.AppSettings["OutputXlsxFolderName"];
            request.OutputXlsxFileName = ConfigurationManager.AppSettings["OutputXlsxFileName"];
            request.OutputXlsxForceAllCellTypeAsString = (ConfigurationManager.AppSettings["OutputXlsxForceAllCellTypeAsString"] == "Yes");

            request.InputCsvFileNameByArgs = null;
            request.OutputXlsxFileNameByArgs = null;

            if (_args.Length >= 1)
            {
                request.InputCsvFileNameByArgs = _args[0];
            }

            if (_args.Length >= 2)
            {
                request.OutputXlsxFileNameByArgs = _args[1];
            }

            MainProcess(request);
        }

        #endregion

        #region リクエスト変数

        /// <summary>
        /// リクエスト変数
        /// </summary>
        public struct Request
        {
            /// <summary>
            /// 入力CSVフォルダ名
            /// </summary>
            public string InputCsvFolderName { get; set; }

            /// <summary>
            /// 入力CSVエンコーディング
            /// </summary>
            public string InputCsvEncoding { get; set; }

            /// <summary>
            /// 入力CSVの最初の一行を読込対象外とするか？
            /// </summary>
            public bool InputCsvSkipFirstRow { get; set; }

            /// <summary>
            /// 出力xlsxフォルダ名
            /// </summary>
            public string OutputXlsxFolderName { get; set; }

            /// <summary>
            /// 出力xlsxファイル名
            /// </summary>
            public string OutputXlsxFileName { get; set; }

            /// <summary>
            /// 出力xlsxのすべてのセルの属性を文字列に設定するか？
            /// </summary>
            public bool OutputXlsxForceAllCellTypeAsString { get; set; }

            /// <summary>
            /// コマンドライン引数で指定された入力CSVファイル名
            /// </summary>
            public string InputCsvFileNameByArgs { get; set; }

            /// <summary>
            /// コマンドライン引数で指定された出力xlsxファイル名
            /// </summary>
            public string OutputXlsxFileNameByArgs { get; set; }
        }

        #endregion

        #region メイン処理

        /// <summary>
        /// メイン処理
        /// </summary>
        private static void MainProcess(Request _request)
        {
            WriteConsoleLogMessage("CSV -> xlsx 変換処理を開始します");

            // 処理する入力ファイルのリストを作成
            // コマンドライン引数で入力ファイル名が指定されていないならば、
            // アプリケーション構成ファイル定義のフォルダ内のすべてのファイルを処理
            // 入力ファイル名がコマンドライン引数で指定されているのならば、そのファイルだけを処理
            FileInfo[] inputFiles;
            if (string.IsNullOrEmpty(_request.InputCsvFileNameByArgs))
            {
                DirectoryInfo di = new DirectoryInfo(_request.InputCsvFolderName);
                inputFiles = di.GetFiles("*.csv", System.IO.SearchOption.TopDirectoryOnly);
                Array.Sort<FileInfo>(inputFiles, delegate (FileInfo a, FileInfo b)
                {
                    return a.Name.CompareTo(b.Name);
                });
            }
            else
            {
                inputFiles = new FileInfo[1] { new FileInfo(_request.InputCsvFileNameByArgs) };
            }

            // 出力ファイルについて
            // 引数で出力ファイル名が指定されていれば、それを使用
            // 指定されていないのならば、アプリケーション構成ファイル定義の値を使用
            string outputXlsxFolderAndFileName;
            outputXlsxFolderAndFileName = _request.OutputXlsxFileNameByArgs ?? Path.Combine(_request.OutputXlsxFolderName, _request.OutputXlsxFileName);

            //  出力xlsxを作成し、シートを追加
            var book = new XLWorkbook();
            var sheet = book.Worksheets.Add(Left(Path.GetFileName(outputXlsxFolderAndFileName), 30));
            int outputRowNumber = 1;

            // 入力ファイルを読み、出力用xlsxに値を設定する
            foreach (var inputFile in inputFiles)
            {
                if (inputFile.Extension.ToLower() != ".csv") continue; // 「.csv2」などの拡張子が検索されるが、それは処理対象外とする
                WriteConsoleLogMessage("読み込んでいます " + inputFile.FullName);

                var parser = new TextFieldParser(inputFile.FullName, Encoding.GetEncoding(_request.InputCsvEncoding));
                using (parser)
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(","); // カンマ区切り

                    parser.HasFieldsEnclosedInQuotes = true; // フィールドが引用符で囲まれているか
                    parser.TrimWhiteSpace = false;           // フィールドの空白トリム設定

                    int inputLineNumber = 0;
                    while (!parser.EndOfData) // ファイルの終端までループ
                    {
                        string[] row = parser.ReadFields(); // フィールドを読込
                        inputLineNumber++;
                        if (_request.InputCsvSkipFirstRow && inputLineNumber == 1) continue; // 最初の一行を読込対象外としている場合、読み飛ばし

                        for (int i = 0; i < row.Length; i++)
                        {
                            var cell = sheet.Cell(outputRowNumber, i + 1);
                            string cellStringValue = Convert.ToString(row[i]);

                            if (_request.OutputXlsxForceAllCellTypeAsString)
                            {
                                cell.SetValue<string>(cellStringValue).Style.NumberFormat.SetFormat("@");
                                continue;
                            }

                            DateTime dateTimeCellValue;
                            if (DateTime.TryParse(cellStringValue, out dateTimeCellValue))
                            {
                                TimeSpan timeSpanCellValue;
                                if (TimeSpan.TryParse(cellStringValue, out timeSpanCellValue))
                                {
                                    cell.SetValue<TimeSpan>(timeSpanCellValue);
                                }
                                else
                                {
                                    cell.SetValue<DateTime>(dateTimeCellValue);
                                }
                                continue;
                            }

                            bool boolCellValue;
                            if (bool.TryParse(cellStringValue, out boolCellValue))
                            {
                                cell.SetValue<bool>(boolCellValue);
                                continue;
                            }

                            decimal decimalCellValue;
                            if (decimal.TryParse(cellStringValue, out decimalCellValue))
                            {
                                cell.SetValue<decimal>(decimalCellValue);
                                continue;
                            }

                            cell.SetValue<string>(cellStringValue);
                        }

                        outputRowNumber++;
                    }
                }
            }

            // 出力用xlsxを書き込む
            WriteConsoleLogMessage("作成しています " + outputXlsxFolderAndFileName);
            book.SaveAs(outputXlsxFolderAndFileName);

            WriteConsoleLogMessage("CSV -> xlsx 変換処理を終了します");
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// コンソールにメッセージを出力します
        /// </summary>
        private static void WriteConsoleLogMessage(string _message)
        {
            Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " " + _message);
        }

        /// <summary>
        /// 文字列の左端から指定された文字数分の文字列を返します
        /// </summary>
        private static string Left(string _target, int _length)
        {
            if (_length <= _target.Length)
            {
                return _target.Substring(0, _length);
            }
            return _target;
        }

        #endregion
    }
}
