using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ProjectX.ExcelParsing
{
    public class EParsing
    {
        public static List<ParsingRow> GetParsingRows(EParsingParam eParam, out List<string> warning)
        {
            warning = new List<string>();
            Stream stream;
            try
            {
                stream = File.Open(eParam.FilePath, FileMode.Open);
            }
            catch (Exception e)
            {
                throw e;
            }

            SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true);
            WorkbookPart workbookPart = document.WorkbookPart;
            WorksheetPart worksheet;

            List<ParsingRow> parsing = new List<ParsingRow>();

            foreach (ESheet param in eParam)
            {
                if (param.SheetName == null || param.SheetName == "")
                {
                    try
                    {
                        worksheet = workbookPart.WorksheetParts.First();
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                }
                else
                {
                    IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == param.SheetName);
                    if (sheets.Count() == 0)
                        throw new ArgumentException("Лист не найден");

                    string relationshipId = sheets.First().Id.Value;
                    worksheet = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                }

                SheetData sheetData = worksheet.Worksheet.Elements<SheetData>().First();

                string[] bufIndex = param.GetBufIndex();

                int countExeption = 0;

                foreach (Row row in sheetData.Elements<Row>())

                {
                    string buf = "";
                    string res = "";

                    try
                    {
                        res = GetCellValue(workbookPart, row, param.PriceIndex);

                        if (!double.TryParse(res, out double price))
                        {
                            continue;
                        }

                        List<ParsingCount> counts = new List<ParsingCount>();

                        foreach (ECount countIndex in param)
                        {
                            try
                            {
                                res = GetCellValue(workbookPart, row, countIndex.Index);
                            }
                            catch
                            {
                                continue;
                            }

                            if (!int.TryParse(res, out int count))
                            {
                                if (!eParam.IsStringValues(res, out count))
                                {
                                    continue;
                                }
                            }

                            counts.Add(new ParsingCount(countIndex.Id, count));
                        }

                        if (counts.Count < 1)
                        {
                            continue;
                        }

                        foreach (string index in bufIndex)
                        {
                            buf += GetCellValue(workbookPart, row, index) + " ";
                        }

                        ParsingRow parsingRow;

                        parsing.Add(parsingRow = new ParsingRow(buf, price, row.RowIndex, eParam.IdProvider));

                        foreach (ParsingCount count in counts)
                        {
                            parsingRow.AddCount(count);
                        }

                        countExeption = 0;
                    }
                    catch
                    {
                        if (countExeption < 100)
                        {
                            countExeption++;
                            continue;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }

            stream.Close();

            warning.ToArray();
            return parsing;
        }

        public static List<ParsingRow> GetParsingRows(EParsingParam[] eParam, out List<string> warning)
        {
            List<string> __warning__ = new List<string>();
            List<ParsingRow> resault = new List<ParsingRow>();
            Parallel.ForEach(eParam, (item) =>
             {
                 resault.InsertRange(resault.Count, GetParsingRows(item, out List<string> warning1));
                 __warning__.InsertRange(__warning__.Count, warning1);
             });
            warning = __warning__;
            return resault;
        }

        private static string GetCellValue(WorkbookPart workbookPart, Row row, string index)
        {
            Cell cell;
            try
            {
                cell = row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, index + row.RowIndex, true) == 0).First();
            }
            catch (Exception e)
            {
                throw e;
            }

            return GetCellValue(cell, workbookPart);
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.DataType == null)
            {
                return cell.InnerText;
            }
            else if (cell.DataType == CellValues.SharedString)
            {
                int id = int.Parse(cell.InnerText);
                return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id).InnerText;
            }
            else
            {
                return cell.InnerText;
            }
        }
    }

    public class EParsingParam : IEnumerable
    {
        public string FilePath { get; set; }
        public string IdProvider { get; set; }
        private List<ESheet> ESheets { get; set; }
        private List<EStringToValue> StringToValues { get; set; }

        public EParsingParam(string filePath, string idProvider)
        {
            ESheets = new List<ESheet>();
            StringToValues = new List<EStringToValue>();
            FilePath = filePath;
            IdProvider = idProvider;
        }

        public void Add(ESheet sheet)
        {
            ESheets.Add(sheet);
        }

        public void AddStringVal(string str, int val)
        {
            StringToValues.Add(new EStringToValue(str, val));
        }

        public bool IsStringValues(string buf, out int val)
        {
            val = -1;
            foreach (var item in StringToValues)
            {
                if (item.IsContain(buf, out val))
                {
                    return true;
                }
            }
            return false;
        }

        public IEnumerator GetEnumerator()
        {
            return ESheets.GetEnumerator();
        }
    }

    public class ESheet : IEnumerable
    {
        public string SheetName { get; set; }
        public string PriceIndex { get; set; }
        private List<string> BufIndex { get; set; }
        private List<ECount> CountIndexes { get; set; }

        public ESheet(string sheetName, string priceIndex)
        {
            BufIndex = new List<string>();
            CountIndexes = new List<ECount>();
            SheetName = sheetName;
            PriceIndex = priceIndex;
        }

        public void AddBufIndex(string value) => BufIndex.Add(value);

        public void AddBufIndex(string[] values) => BufIndex.AddRange(values);

        public string[] GetBufIndex() => BufIndex.ToArray();

        public void AddCountIndex(string id, string value) => CountIndexes.Add(new ECount(id, value));

        public IEnumerator GetEnumerator()
        {
            return CountIndexes.GetEnumerator();
        }
    }

    public class EStringToValue
    {
        public string Str { get; set; }
        public int Value { get; set; }

        public EStringToValue(string str, int value)
        {
            Str = str;
            Value = value;
        }

        public bool IsContain(string buf, out int val)
        {
            if (Str.ToLower().Equals(buf.Trim().ToLower()))
            {
                val = Value;
                return true;
            }
            else
            {
                val = -1;
                return false;
            }
        }
    }

    public class ECount
    {
        public string Id { get; private set; }
        public string Index { get; private set; }

        public ECount(string id, string index)
        {
            Id = id;
            Index = index;
        }
    }
}