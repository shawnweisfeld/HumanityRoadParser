using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace HumanityRoadParser
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo source = new FileInfo("source.xlsx");

            if (!source.Exists)
            {
                Console.WriteLine("I could not find that file, please try again");
            }
            else
            {
                Console.WriteLine("Parsing Excel File '{0}'.", source.FullName);
                var infos = ParseExcelFile(source);
                WriteNewExcelFile(infos, source.FullName);
            }

            Console.WriteLine("Done!");
        }

        private static void WriteNewExcelFile(List<InfoObject> infos, string path)
        {
            string dir = Path.GetDirectoryName(path);
            string fileName = Path.GetFileNameWithoutExtension(path) + " Parsed";
            string fileExt = Path.GetExtension(path);
            string newFileName = Path.Combine(dir, string.Format("{0}{1}", fileName, fileExt));
            int revision = 0;

            while (File.Exists(newFileName))
            {
                revision++;
                newFileName = Path.Combine(dir, string.Format("{0} ({1}){2}", fileName, revision, fileExt));
            }

            Console.WriteLine("Exporting Results {0}", newFileName);

            using (var package = new ExcelPackage(new FileInfo(newFileName)))
            {
                int row = 1;
                int col = 1;

                //print headings
                var worksheet = package.Workbook.Worksheets.Add("Results");
                worksheet.Cells[row, col++].Value = "State";
                worksheet.Cells[row, col++].Value = "Row";
                worksheet.Cells[row, col++].Value = "Column";
                worksheet.Cells[row, col++].Value = "AgencyType";
                worksheet.Cells[row, col++].Value = "Name";
                worksheet.Cells[row, col++].Value = "ColumnName";
                worksheet.Cells[row, col++].Value = "ValueType";
                worksheet.Cells[row, col++].Value = "Value";

                //print data
                foreach (var info in infos.OrderBy(x => x.State).ThenBy(x => x.Row).ThenBy(x => x.Column))
                {
                    col = 1;
                    row++;
                    worksheet.Cells[row, col++].Value = info.State;
                    worksheet.Cells[row, col++].Value = info.Row;
                    worksheet.Cells[row, col++].Value = info.Column;
                    worksheet.Cells[row, col++].Value = info.AgencyType;
                    worksheet.Cells[row, col++].Value = info.Name;
                    worksheet.Cells[row, col++].Value = info.ColumnName;
                    worksheet.Cells[row, col++].Value = info.ValueType;
                    worksheet.Cells[row, col++].Value = info.Value;
                }

                worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].AutoFitColumns();

                package.Save();
            }

            System.Diagnostics.Process.Start(newFileName);
        }

        static Regex _uriRegex = new Regex(@"(mailto\:|(news|(ht|f)tp(s?))\://)(([^[:space:]]+)|([^[:space:]]+)( #([^#]+)#)?)");
        static Regex _twitterRegex = new Regex(@"^@([a-zA-Z0-9_]{1,15})$");

        private static ValueType GetValueType(string value)
        {
            if (value.StartsWith("https://www.facebook.com", StringComparison.InvariantCultureIgnoreCase)
                || value.StartsWith("http://www.facebook.com", StringComparison.InvariantCultureIgnoreCase))
                return ValueType.Facebook;
            else if (_uriRegex.IsMatch(value))
                return ValueType.Url;
            else if (_twitterRegex.IsMatch(value))
                return ValueType.Twitter;
            else
                return ValueType.String;
        }


        private static List<InfoObject> ParseExcelFile(FileInfo fileInfo)
        {
            var infos = new List<InfoObject>();

            using (var package = new ExcelPackage(fileInfo))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Name.Length > 3)
                    {
                        Console.WriteLine("Skipping Worksheet '{0}' - it is not a state sheet.", worksheet.Name);
                    }
                    else
                    {
                        Console.WriteLine("Parsing Worksheet '{0}'.", worksheet.Name);
                        infos.AddRange(ParseSheet(worksheet));
                    }
                }
            }

            return infos;
        }
        private static IEnumerable<InfoObject> ParseSheet(ExcelWorksheet worksheet)
        {
            int row = 1;
            string agencyType = string.Empty;
            Dictionary<int, string> columnNames = new Dictionary<int, string>();

            while (row <= worksheet.Dimension.End.Row)
            {
                //header row
                if (worksheet.Cells[row, 1].Style.Fill.BackgroundColor.Rgb == "FFFFFF00")
                {
                    agencyType = worksheet.Cells[row, 1].GetValue<string>();

                    for (int i = 2; i <= worksheet.Dimension.End.Column; i++)
                    {
                        var columnName = worksheet.Cells[row, i].GetValue<string>();
                        if (!string.IsNullOrEmpty(columnName))
                        {
                            if (columnNames.ContainsKey(i))
                                columnNames[i] = columnName;
                            else
                                columnNames.Add(i, columnName);
                        }
                    }
                }
                else
                {
                    string name = worksheet.Cells[row, 1].GetValue<string>();

                    if (string.IsNullOrEmpty(name)
                        || name.StartsWith("Note:", StringComparison.InvariantCultureIgnoreCase)
                        || worksheet.Cells[row, 1].Style.Font.Italic)
                    {
                        //skip these rows
                    }
                    else
                    {

                        for (int i = 2; i <= worksheet.Dimension.End.Column; i++)
                        {
                            var columnName = string.Empty;
                            columnNames.TryGetValue(i, out columnName);

                            if (string.IsNullOrEmpty(columnName))
                                columnName = "Unknown Column";

                            var obj = new InfoObject()
                            {
                                Row = row,
                                Column = i,
                                State = worksheet.Name.Trim(),
                                AgencyType = agencyType,
                                Name = name.Trim(),
                                ColumnName = columnName,
                                Value = worksheet.Cells[row, i].GetValue<string>()
                            };

                            if (!string.IsNullOrEmpty(obj.Value))
                            {
                                obj.ValueType = GetValueType(obj.Value);
                                yield return obj;
                            }
                        }
                    }
                }

                row++;
            }
        }

    }
}
