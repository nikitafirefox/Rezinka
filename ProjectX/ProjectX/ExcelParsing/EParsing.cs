using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ProjectX.ExcelParsing
{
    public class EParsing
    {
        public static int MaxCount { get; private set;}

        public static string FilePath { get; private set;}

        public static int CurentIndex { get; set;}

        public static List<ParsingRow> GetParsingRows(EParsingParam eParam, out List<WarningParsingRow> warning)
        {
            warning = new List<WarningParsingRow>();
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

                FilePath = param.SheetName ?? "[Default]";

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

                MaxCount = sheetData.Elements<Row>().Count();

                foreach (Row row in sheetData.Elements<Row>())
                {

                    CurentIndex = int.Parse(row.RowIndex);

                    string buf = "";
                    string res = "";

                    string sheetName = param.SheetName ?? "[Default]";

                    try
                    {
                        res = GetCellValue(workbookPart, row, param.PriceIndex);

                        if (res != "" && res != null) {
                            res = param.Replace(res, param.PriceIndex); 
                        }

                        if (!double.TryParse(res, out double price))
                        {
                            warning.Add(new WarningParsingRow()
                            {
                                FileName = eParam.FilePath,
                                SheetName = sheetName,
                                RowIndex = row.RowIndex,
                                CellIndex = param.PriceIndex,
                                Message = "Не распознана цена"
                            });
                            continue;
                        }

                        List<ParsingCount> counts = new List<ParsingCount>();

                        string countIndexes = "";

                        foreach (ECount countIndex in param)
                        {


                          

                            countIndexes += countIndex.Index + ",";

                            try
                            {
      

                                res = GetCellValue(workbookPart, row, countIndex.Index);

                                if (res != "" && res!=null)
                                {
                                    res = param.Replace(res, countIndex.Index);
                                }

                            }
                            catch
                            {

                                continue;
                            }

                            if (!int.TryParse(res, out int count))
                            {
                                if (!eParam.IsStringValues(res, out count))
                                {
                                    if (res.Trim() != "")
                                    {

                                        warning.Add(new WarningParsingRow()
                                        {
                                            FileName = eParam.FilePath,
                                            SheetName = sheetName,
                                            RowIndex = row.RowIndex,
                                            CellIndex = countIndex.Index,
                                            Message = "Не распознан остаток"
                                        });
                                    }

                                    continue;
                                }
                            }

                            counts.Add(new ParsingCount(countIndex.Id, count));
                        }

                        if (counts.Count < 1)
                        {

                            warning.Add(new WarningParsingRow()
                            {
                                FileName = eParam.FilePath,
                                SheetName = sheetName,
                                RowIndex = row.RowIndex,
                                CellIndex = countIndexes.TrimEnd(','),
                                Message = "Отсутствует остаток"
                            });


                            continue;
                        }

                        foreach (string index in bufIndex)
                        {
                            string promBuffer;


                            Cell cell = row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, index + row.RowIndex, true) == 0).FirstOrDefault();

                            if (cell == null)
                            {
                                promBuffer = "";
                            }
                            else
                            {
                                promBuffer = param.Replace(GetCellValue(cell,workbookPart), index);
                            }



                            buf += promBuffer + " ";
                        }

                        ParsingRow parsingRow;

                        parsing.Add(parsingRow = new ParsingRow(buf, price, row.RowIndex, eParam.IdProvider) {
                            FileName = eParam.FilePath,
                            SheetName = param.SheetName??"[Default]"
                        });

                        foreach (ParsingCount count in counts)
                        {
                            parsingRow.AddCount(count);
                        }

                        countExeption = 0;
                    }
                    catch
                    {
                        warning.Add(new WarningParsingRow()
                        {
                            FileName = eParam.FilePath,
                            SheetName = sheetName,
                            RowIndex = row.RowIndex,
                            CellIndex = param.PriceIndex,
                            Message = "Отсутствует цена"
                        });

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

        public static List<ParsingRow> GetParsingRows(EParsingParam[] eParam, out List<WarningParsingRow> warning)
        {
            List<WarningParsingRow> __warning__ = new List<WarningParsingRow>();
            List<ParsingRow> resault = new List<ParsingRow>();

            foreach (EParsingParam param in eParam) {

                
                MaxCount = int.MaxValue;
                CurentIndex = 0;

                resault.AddRange(GetParsingRows(param, out List<WarningParsingRow> warning1));
                __warning__.AddRange(warning1);
            }

            /*
            Parallel.ForEach(eParam, (item) =>
             {
                 resault.InsertRange(resault.Count, GetParsingRows(item, out List<string> warning1));
                 __warning__.InsertRange(__warning__.Count, warning1);
             });
            */

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


    [Serializable]
    public class EParsingParam : IEnumerable
    {
        public string FilePath { get; set; }
        public string IdProvider { get; set; }
        private List<ESheet> ESheets { get; set; }
        private List<EStringToValue> StringToValues { get; set; }

        public EParsingParam() {
            ESheets = new List<ESheet>();
            StringToValues = new List<EStringToValue>();
            FilePath = "";
            IdProvider = null;
        }

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
            if (StringToValues.FindIndex(x => x.Str == str) == -1)
            {
                StringToValues.Add(new EStringToValue(str, val));
            }
        }

        public void DeleteStringVal(string str) {
            var item = StringToValues.Find(x => x.Str == str);
            StringToValues.Remove(item);
        }

        public void ChangeStringVal(string str1, string str2, int val) {
            var item = StringToValues.Find(x => x.Str == str1);
            item.Str = str2;
            item.Value = val;
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

        public EStringToValue GetEStringToValue(string str) {
            EStringToValue value = null;
            value = StringToValues.Find(x => x.Str == str);

            return value;
        }

        public IEnumerator GetEnumerator()
        {
            return ESheets.GetEnumerator();
        }

        public IEnumerable<EStringToValue> GetEStringToValues() {
            return StringToValues;
        }

        public ESheet SheetByName(string name) {
            return ESheets.Find(x => x.SheetName == name);
        }

        public void DeleteSheet(string name) {
            ESheet sheet = SheetByName(name);
            ESheets.Remove(sheet);
        }

        public void ChangeSheet(string name, ESheet sheet) {

            DeleteSheet(name);
            Add(sheet);

        }

        public void ClearSheets() => ESheets.Clear();


    }

    [Serializable]
    public class ESheet : IEnumerable
    {
        public string SheetName { get; set; }
        public string PriceIndex { get; set; }
        private List<string> BufIndex { get; set; }
        private List<ECount> CountIndexes { get; set; }
        private List<EReplacementCell> EReplacementCells { get; set; }

        public ESheet() {

            SheetName = null;
            PriceIndex = "";
            BufIndex = new List<string>();
            CountIndexes = new List<ECount>();
            EReplacementCells = new List<EReplacementCell>();
        }

        public ESheet(string sheetName, string priceIndex)
        {
            BufIndex = new List<string>();
            CountIndexes = new List<ECount>();
            SheetName = sheetName;
            PriceIndex = priceIndex;
            EReplacementCells = new List<EReplacementCell>();
        }

        public void AddBufIndex(string value) {

            if (!BufIndex.Contains(value))
            {
                BufIndex.Add(value);
            }

        }

        public void AddBufIndex(string[] values) => BufIndex.AddRange(values);

        public void DeleteBufIndex(string str) {
            BufIndex.Remove(str);
        }

        public void ChangeBufIndex(string str1, string str2) {
            DeleteBufIndex(str1);
            AddBufIndex(str2);
        }

        public void AddReplacementCell(EReplacementCell cell) => EReplacementCells.Add(cell);

        public string[] GetBufIndex() => BufIndex.ToArray();

        public string Replace(string str, string cellIndex) {
            if (EReplacementCells.Count > 0)
            {
                string outStr = str;
                foreach (var item in EReplacementCells)
                {
                    if (item.CellIndex.Equals(cellIndex)) {
                        outStr = item.Replace(str);
                    }
                }
                return outStr;
            }
            else {
                return str;
            }
        }

        public void AddCountIndex(string id, string value) => CountIndexes.Add(new ECount(id, value));

        public void DeleteCountIndex(string id) {
            ECount count = CountIndexes.Find(x => x.Id == id);
            CountIndexes.Remove(count);
        }

        public void ChangeCountIndex(string id1, string id2, string value) {

            DeleteCountIndex(id1);
            AddCountIndex(id2, value);
        }

        public EReplacementCell GetReplacementCell(string index) =>
            EReplacementCells.Find(x => x.CellIndex == index);

        public void DeleteReplacementCell(string index) {
            EReplacementCell cell = GetReplacementCell(index);

            if (cell != null) {
                EReplacementCells.Remove(cell);
            }
        }

        public IEnumerator GetEnumerator()
        {
            return CountIndexes.GetEnumerator();
        }

        public IEnumerable GetBufIndexList() {
            return BufIndex;
        }

        public IEnumerable GetReplacementCells() {
            return EReplacementCells;
        }
    }

    [Serializable]
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

    [Serializable]
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

    [Serializable]
    public class EReplacementCell:IEnumerable {

        public string CellIndex { get; set;}

        private List<EReplacementString> EReplacementStrings { get; set; }

        public EReplacementCell()
        {
            EReplacementStrings = new List<EReplacementString>();
        }

        public EReplacementCell(string cellIndex) {
            CellIndex = cellIndex;
            EReplacementStrings = new List<EReplacementString>();
        }

        public void Add(string str, string repStr) {
            EReplacementStrings.Add(new EReplacementString(str, repStr));
        }

        public void DeleteStr(string str) {
            EReplacementStrings.Remove(EReplacementStrings.Find(x => x.Str == str));
        }

        public void ChangeStr(string strOld,string strNew, string repStr) {
            DeleteStr(strOld);
            Add(strNew, repStr);
        }

        public string Replace(string str) {
            string outStr = str;
            foreach (var item in EReplacementStrings) {
                outStr = item.Replace(str);
            }
            return outStr;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)EReplacementStrings).GetEnumerator();
        }
    }

    [Serializable]
    public class EReplacementString {
        public string Str { get; set;}
        public string ReplceStr { get; set;}

        public EReplacementString(string str, string replaceStr) {
            Str = str;
            ReplceStr = replaceStr;
        }

        public string Replace(string str) {
            return str.Equals(Str) ? ReplceStr : str;
        }

    }
}