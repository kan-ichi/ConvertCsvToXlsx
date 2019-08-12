using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.IO;
using ClosedXML.Excel;
using System.Text;

namespace ConvertCsvToXlsxTest
{
    [TestClass]
    public class ProgramTest
    {
        /// <summary>
        /// メイン処理のテスト
        /// ・入力CSVのファイル名が指定されていない場合、入力CSVのフォルダのすべてのファイルが処理対象となる
        /// ・入力ファイルの拡張子が「csv」でない場合、読み飛ばされる
        /// ・出力xlsxのすべてのセルの属性を文字列に設定する
        /// ・入力CSVと出力xlsxの内容が一致する
        /// </summary>
        [TestMethod]
        public void TestMethod1()
        {
            string testMethodName = MethodBase.GetCurrentMethod().Name;
            string testingFolderName = Path.Combine(GetAppPath(), testMethodName);
            Directory.CreateDirectory(testingFolderName);
            Encoding enc = Encoding.GetEncoding("Shift_JIS");

            string inputCsvPathX;
            {
                string inputCsvFileName = testMethodName + ".csvX";
                inputCsvPathX = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] { };
                File.WriteAllLines(inputCsvPathX, lines, enc);
            }

            string inputCsvPath;
            {
                string inputCsvFileName = testMethodName + ".CSV";
                inputCsvPath = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] {
                    "1a,\"1,b\",1c",
                    "2a,,2c,2d",
                    "3a, "
                };
                File.WriteAllLines(inputCsvPath, lines, enc);
            }

            string outputXlsxPath;
            {
                string outputXlsxFileName = testMethodName + ".xlsx";
                outputXlsxPath = Path.Combine(testingFolderName, outputXlsxFileName);
            }

            ConvertCsvToXlsx.Program.Request request = new ConvertCsvToXlsx.Program.Request();
            {
                request.InputCsvFolderName = testingFolderName;
                request.InputCsvEncoding = enc.WebName;
                request.OutputXlsxFileNameByArgs = outputXlsxPath;
                request.OutputXlsxForceAllCellTypeAsString = true;
            }

            {
                var target = new PrivateType(typeof(ConvertCsvToXlsx.Program));
                target.InvokeStatic("MainProcess", request);
            }

            using (var workbook = new XLWorkbook(outputXlsxPath))
            {
                string sheetName = Path.GetFileName(outputXlsxPath);
                var worksheet = workbook.Worksheet(sheetName);
                var outputXlsxTable = worksheet.RangeUsed().AsTable();

                Assert.AreEqual(3, outputXlsxTable.RowCount());
                Assert.AreEqual("1a", worksheet.Cell(1, 1).Value);
                Assert.AreEqual("1,b", worksheet.Cell(1, 2).Value);
                Assert.AreEqual("1c", worksheet.Cell(1, 3).Value);
                Assert.AreEqual("2a", worksheet.Cell(2, 1).Value);
                Assert.AreEqual("", worksheet.Cell(2, 2).Value);
                Assert.AreEqual("2c", worksheet.Cell(2, 3).Value);
                Assert.AreEqual("3a", worksheet.Cell(3, 1).Value);
                Assert.AreEqual(" ", worksheet.Cell(3, 2).Value);
                Assert.AreEqual("", worksheet.Cell(3, 3).Value);
            }

            File.Delete(inputCsvPathX);
            File.Delete(inputCsvPath);
            File.Delete(outputXlsxPath);
            Directory.Delete(testingFolderName);
        }

        /// <summary>
        /// エントリーポイントからメイン処理のテスト
        /// ・入力CSVのレコードがない場合、出力xlsxのセルには何も出力されない（使用中のセル範囲がnullとなる）
        /// ・出力xlsxのシート名は出力xlsxのファイル名であるが、先頭30文字のみ使用される
        /// </summary>
        [TestMethod]
        public void TestMethod2()
        {
            string testMethodName = MethodBase.GetCurrentMethod().Name;
            string testingFolderName = GetAppPath();
            Encoding enc = Encoding.GetEncoding("Shift_JIS");

            string inputCsvPath;
            {
                string inputCsvFileName = testMethodName + ".csv";
                inputCsvPath = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] { };
                File.WriteAllLines(inputCsvPath, lines, enc);
            }

            string outputXlsxFileName = testMethodName + "123456789012345678901234567890" + ".xlsx";
            string outputXlsxPath;
            {
                outputXlsxPath = Path.Combine(testingFolderName, outputXlsxFileName);
            }

            {
                string[] args = new string[] { inputCsvPath, outputXlsxPath };
                ConvertCsvToXlsx.Program.Main(args);
            }

            using (var workbook = new XLWorkbook(outputXlsxPath))
            {
                string sheetName = outputXlsxFileName.Substring(0, 30);
                var worksheet = workbook.Worksheet(sheetName);
                Assert.IsNull(worksheet.RangeUsed());
            }

            File.Delete(inputCsvPath);
            File.Delete(outputXlsxPath);
        }

        /// <summary>
        /// エントリーポイントのテスト
        /// ・入力ファイル名がブランクであるため、例外発生
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(System.ArgumentException))]
        public void TestMethod3()
        {
            string[] args = new string[] { };
            ConvertCsvToXlsx.Program.Main(args);
        }

        /// <summary>
        /// メイン処理のテスト
        /// ・入力CSVの最初の一行を読込対象外とした場合、最初の一行は読込対象外となる
        /// </summary>
        [TestMethod]
        public void TestMethod4()
        {
            string testMethodName = MethodBase.GetCurrentMethod().Name;
            string testingFolderName = GetAppPath();
            Encoding enc = Encoding.GetEncoding("utf-8");

            string inputCsvPath;
            {
                string inputCsvFileName = testMethodName + ".CSV";
                inputCsvPath = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] {
                    "1a,\"1,b\",1c",
                    "2a,,2c,2d",
                    "3a, "
                };
                File.WriteAllLines(inputCsvPath, lines, enc);
            }

            string outputXlsxPath;
            {
                string outputXlsxFileName = testMethodName + ".xlsx";
                outputXlsxPath = Path.Combine(testingFolderName, outputXlsxFileName);
            }

            ConvertCsvToXlsx.Program.Request request = new ConvertCsvToXlsx.Program.Request();
            {
                request.InputCsvFileNameByArgs = inputCsvPath;
                request.InputCsvEncoding = enc.WebName;
                request.InputCsvSkipFirstRow = true;
                request.OutputXlsxFileNameByArgs = outputXlsxPath;
            }

            {
                var target = new PrivateType(typeof(ConvertCsvToXlsx.Program));
                target.InvokeStatic("MainProcess", request);
            }

            using (var workbook = new XLWorkbook(outputXlsxPath))
            {
                string sheetName = Path.GetFileName(outputXlsxPath);
                var worksheet = workbook.Worksheet(sheetName);
                var outputXlsxTable = worksheet.RangeUsed().AsTable();

                Assert.AreEqual(2, outputXlsxTable.RowCount());
                Assert.AreEqual("2a", worksheet.Cell(1, 1).Value);
                Assert.AreEqual("", worksheet.Cell(1, 2).Value);
                Assert.AreEqual("2c", worksheet.Cell(1, 3).Value);
                Assert.AreEqual("3a", worksheet.Cell(2, 1).Value);
                Assert.AreEqual(" ", worksheet.Cell(2, 2).Value);
                Assert.AreEqual("", worksheet.Cell(2, 3).Value);
            }

            File.Delete(inputCsvPath);
            File.Delete(outputXlsxPath);
        }

        /// <summary>
        /// メイン処理のテスト
        /// ・CSVの文字列をデータ変換する処理のテスト
        /// ・時刻 -> 日付時刻 -> 日付 -> 真偽値 -> 数値 -> 文字列 の順に判定する
        /// </summary>
        [TestMethod]
        public void TestMethod5()
        {
            string testMethodName = MethodBase.GetCurrentMethod().Name;
            string testingFolderName = GetAppPath();
            Encoding enc = Encoding.GetEncoding("Shift_JIS");

            string inputCsvPath;
            {
                string inputCsvFileName = testMethodName + ".CSV";
                inputCsvPath = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] {
                    "12:34:56           , 0:0         , 23:59:59           , -00:00, 24:00",
                    "2000/02/29         , 1900/1/1    , 2999/12/31         , 2000-01-01, 1999/02/29",
                    "2000/02/29 12:34:56, 1900/1/1 0:0, 2999/12/31 23:59:59, 2000/01/01 00:00:60",
                    "True               , false       , TRUE               , falsE, TRUE/FALSE",
                    "-1.2345678901234567, -0          , 123456789012345678 , １２３４５６７８"
                };
                File.WriteAllLines(inputCsvPath, lines, enc);
            }

            string outputXlsxPath;
            {
                string outputXlsxFileName = testMethodName + ".xlsx";
                outputXlsxPath = Path.Combine(testingFolderName, outputXlsxFileName);
            }

            ConvertCsvToXlsx.Program.Request request = new ConvertCsvToXlsx.Program.Request();
            {
                request.InputCsvFileNameByArgs = inputCsvPath;
                request.InputCsvEncoding = enc.WebName;
                request.OutputXlsxFileNameByArgs = outputXlsxPath;
            }

            {
                var target = new PrivateType(typeof(ConvertCsvToXlsx.Program));
                target.InvokeStatic("MainProcess", request);
            }

            using (var workbook = new XLWorkbook(outputXlsxPath))
            {
                string sheetName = Path.GetFileName(outputXlsxPath);
                var worksheet = workbook.Worksheet(sheetName);
                var outputXlsxTable = worksheet.RangeUsed().AsTable();

                Assert.AreEqual(new TimeSpan(12, 34, 56), TimeSpan.FromDays(worksheet.Cell(1, 1).GetValue<double>()));
                Assert.AreEqual(new TimeSpan(00, 00, 00), TimeSpan.FromDays(worksheet.Cell(1, 2).GetValue<double>()));
                Assert.AreEqual(new TimeSpan(23, 59, 59), TimeSpan.FromDays(worksheet.Cell(1, 3).GetValue<double>()));
                Assert.AreEqual(" -00:00", worksheet.Cell(1, 4).GetValue<string>());
                Assert.AreEqual(" 24:00", worksheet.Cell(1, 5).GetValue<string>());

                Assert.AreEqual(new DateTime(2000, 02, 29), worksheet.Cell(2, 1).GetValue<DateTime>());
                Assert.AreEqual(new DateTime(1900, 01, 01), worksheet.Cell(2, 2).GetValue<DateTime>());
                Assert.AreEqual(new DateTime(2999, 12, 31), worksheet.Cell(2, 3).GetValue<DateTime>());
                Assert.AreEqual(new DateTime(2000, 01, 01), worksheet.Cell(2, 4).GetValue<DateTime>());
                Assert.AreEqual(" 1999/02/29", worksheet.Cell(2, 5).GetValue<string>());

                Assert.AreEqual(new DateTime(2000, 02, 29, 12, 34, 56), worksheet.Cell(3, 1).GetValue<DateTime>());
                Assert.AreEqual(new DateTime(1900, 01, 01, 00, 00, 00), worksheet.Cell(3, 2).GetValue<DateTime>());
                Assert.AreEqual(new DateTime(2999, 12, 31, 23, 59, 59), worksheet.Cell(3, 3).GetValue<DateTime>());
                Assert.AreEqual(" 2000/01/01 00:00:60", worksheet.Cell(3, 4).GetValue<string>());

                Assert.AreEqual(true, worksheet.Cell(4, 1).GetValue<bool>());
                Assert.AreEqual(false, worksheet.Cell(4, 2).GetValue<bool>());
                Assert.AreEqual(true, worksheet.Cell(4, 3).GetValue<bool>());
                Assert.AreEqual(false, worksheet.Cell(4, 4).GetValue<bool>());
                Assert.AreEqual(" TRUE/FALSE", worksheet.Cell(4, 5).GetValue<string>());

                Assert.AreEqual(-1.2345678901234600m, worksheet.Cell(5, 1).GetValue<decimal>()); // Excelの仕様で15桁目以降に誤差が生じる
                Assert.AreEqual(0m, worksheet.Cell(5, 2).GetValue<decimal>());
                Assert.AreEqual(123456789012346000m, worksheet.Cell(5, 3).GetValue<decimal>()); // Excelの仕様で15桁目以降に誤差が生じる
                Assert.AreEqual(" １２３４５６７８", worksheet.Cell(5, 4).GetValue<string>());
            }

            File.Delete(inputCsvPath);
            File.Delete(outputXlsxPath);
        }

        /// <summary>
        /// テスト実施中のフォルダパスを取得します
        /// </summary>
        private static string GetAppPath()
        {
            string path = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase;
            //URIを通常のパス形式に変換する
            Uri u = new Uri(path);
            path = u.LocalPath + Uri.UnescapeDataString(u.Fragment);
            return System.IO.Path.GetDirectoryName(path);
        }

    }
}
