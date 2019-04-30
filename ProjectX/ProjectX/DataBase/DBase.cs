using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ProjectX.Dict;
using ProjectX.ExcelParsing;
using ProjectX.Information;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace ProjectX.DataBase
{
    public class DBase : IEnumerable
    {
        public IEnumerator GetEnumerator()
        {
            return DBRows.GetEnumerator();
        }

        private GenId GenId { get; set; }
        public double DefaultMarkup { get; set; }
        private readonly string pathXML;
        private List<DBRow> DBRows { get; set; }

        public DBase()
        {
            DBRows = new List<DBRow>();
            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\DataBase.xml");

            if (!File.Exists(pathXML))
            {
                string dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data");
                if (!Directory.Exists(dataDirectory)) { Directory.CreateDirectory(dataDirectory); }
                GenId = new GenId('A', -1, 1);
                DefaultMarkup = 10;
                FileStream fs = new FileStream(pathXML, FileMode.Create);
                XmlTextWriter xmlOut = new XmlTextWriter(fs, Encoding.Unicode)
                {
                    Formatting = Formatting.Indented
                };
                xmlOut.WriteStartDocument();
                xmlOut.WriteStartElement("root");
                xmlOut.WriteEndElement();
                xmlOut.WriteEndDocument();
                xmlOut.Close();
                fs.Close();
                Save();
            }
            else
            {
                InitBase();
            }
        }

        private void InitBase()
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;

            XmlNode xmlSet = xroot.SelectSingleNode("settings");
            GenId = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            DefaultMarkup = double.Parse(xmlSet.ChildNodes.Item(3).InnerText);

            XmlNode xmlNode = xroot.GetElementsByTagName("dRows").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                DBRows.Add(new DBRow(x));
            }
        }

        public void AddRows(ParsingRow[] parsingRows)
        {
            string[] idProv = parsingRows.Select(x => x.IdProvider).Distinct().ToArray();

            DBRows.RemoveAll(x => idProv.Contains(x.IdPosition.Split('-').First()));

            foreach (var item in parsingRows)
            {
                GResault gResault = (GResault)item.Resault;
                foreach (ParsingCount item1 in item)
                {
                    AddRow(gResault.Id, gResault.Name, item.IdProvider + '-' + item1.Id, item.Price, DefaultMarkup, item1.Count, gResault.Addition);
                }
            }
        }

        public void Save()
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;
            xroot.RemoveAll();

            xroot.AppendChild(GenId.GetXmlNode(xmlDocument));
            XmlNode xmlSet = xroot.SelectSingleNode("settings");
            XmlElement element = xmlDocument.CreateElement("defaultMarkup");
            element.InnerText = DefaultMarkup.ToString();
            xmlSet.AppendChild(element);

            XmlElement rowsElement;
            xroot.AppendChild(rowsElement = xmlDocument.CreateElement("dRows"));
            foreach (var item in DBRows)
            {
                rowsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }
            xmlDocument.Save(pathXML);
        }

        private void AddRow(string idProduct, string nameProduct, string idPosition, double priceProv, double markup,
            int count,string addition)
        {
            DBRows.Add(new DBRow(GenId.NexVal(), idProduct, nameProduct, idPosition, priceProv, markup, count,addition));
        }

        private void CellSave(SharedStringTablePart shareStringPart, Row row, Cell cell, string value, string cellReference, int index)
        {
            shareStringPart.SharedStringTable.AppendChild(
                         new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(value)));
            cell = new Cell() { CellReference = cellReference };
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            row.AppendChild(cell);
        }

        private void CellSave(SharedStringTablePart shareStringPart, Row row, Cell cell, string value, string cellReference)
        {
            CellSave(shareStringPart, row, cell, value, cellReference, shareStringPart.SharedStringTable.Count());
        }

        private void CreateExcel(string filepath)
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            SheetData sheetData;
            worksheetPart.Worksheet = new Worksheet(sheetData = new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild(new Sheets());

            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Data"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            spreadsheetDocument.Close();
        }

        public void SaveDataExcel(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics, ExcelProviderOutParametrics providerOutParametrics, ExcelProductOutParametrics productOutParametrics)
        {
            CreateExcel(filepath);

            ExcelCellList cellList = new ExcelCellList();

            var res = DBRows.Where(x => idProviders.Contains(x.IdPosition.Split('-').First()));

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true))
            {
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                if (shareStringPart.SharedStringTable == null)
                {
                    shareStringPart.SharedStringTable = new SharedStringTable();
                }

                int shareIndex = 0;

                WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                WorksheetPart worksheet;
                IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == "Data");
                if (sheets.Count() == 0)
                    throw new ArgumentException("Лист не найден");
                string relationshipId = sheets.First().Id.Value;
                worksheet = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheet.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                Cell cell = null;

                Row row = new Row() { RowIndex = 1 };

                if (defaultOutParametrics.IsIdRows) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_IdRows, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsIdProduct) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_IdProduct, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsNameProduct) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_NameProduct, cellList.NextVal() + "1", shareIndex++);


                if (productOutParametrics.IsBrandName) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandName, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsBrandCountry) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandCountry, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsBrandDescription) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandDescription, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsBrandRunFlatName) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandRunFlatName, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelName) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelName, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelType) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelType, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelSeason) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelSeason, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelCommercial) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelCommercial, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelDescription) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelDescription, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelWhileLetters) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelWhileLetters, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelMudSnow) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelMudSnow, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingWidth) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingWidth, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingHeight) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingHeight, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingDiameter) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingDiameter, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingLoadIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingLoadIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingSpeedIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingSpeedIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingSpikes) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingSpikes, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingCountry) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingCountry, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingTemperatureIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingTemperatureIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingTractionIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingTractionIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingTreadwearIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingTreadwearIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingExtraLoad) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingExtraLoad, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingFlangeProtection) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingFlangeProtection, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingRunFlat) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingRunFlat, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingAccomadation) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingAccomadation, cellList.NextVal() + "1", shareIndex++);
                


                if (defaultOutParametrics.IsIdPosition) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_IdPosition, cellList.NextVal() + "1", shareIndex++);
                if (providerOutParametrics.IsProviderName) CellSave(shareStringPart, row, cell, providerOutParametrics.Name_ProviderName, cellList.NextVal() + "1", shareIndex++);
                if (providerOutParametrics.IsProviderPriority) CellSave(shareStringPart, row, cell, providerOutParametrics.Name_ProviderPriority, cellList.NextVal() + "1", shareIndex++);
                if (providerOutParametrics.IsStockName) CellSave(shareStringPart, row, cell, providerOutParametrics.Name_StockName, cellList.NextVal() + "1", shareIndex++);
                if (providerOutParametrics.IsStockTime) CellSave(shareStringPart, row, cell, providerOutParametrics.Name_StockTime, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsProviderPrice) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_ProviderPrice, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsMarkup) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_Markup, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsTotalPrice) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_TotalPrice, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsCount) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_Count, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsAddition) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_Addition, cellList.NextVal() + "1", shareIndex++);

                sheetData.Append(row);

                Dictionary<string, object> valuePairsProvider = null;

                Dictionary<string, object> valuePairsDictionary = null;

                uint index = 2;
                foreach (var item in res)
                {
                    cellList.Restart();

                    row = new Row() { RowIndex = index };

                    if (providerOutParametrics.IsProviderName || providerOutParametrics.IsProviderPriority || providerOutParametrics.IsStockName || providerOutParametrics.IsStockTime)
                    {
                        valuePairsProvider = providerOutParametrics.ProvidersSrc.GetValuesById(item.IdPosition.Split('-').First(), item.IdPosition.Split('-').Last());
                    }
                    if (productOutParametrics.IsBrandName || productOutParametrics.IsBrandCountry || productOutParametrics.IsBrandDescription || productOutParametrics.IsBrandRunFlatName ||
                        productOutParametrics.IsModelName || productOutParametrics.IsModelType || productOutParametrics.IsModelSeason || productOutParametrics.IsModelCommercial ||
                        productOutParametrics.IsModelDescription || productOutParametrics.IsModelWhileLetters || productOutParametrics.IsModelMudSnow ||
                        productOutParametrics.IsMarkingWidth || productOutParametrics.IsMarkingHeight || productOutParametrics.IsMarkingDiameter || productOutParametrics.IsMarkingLoadIndex ||
                        productOutParametrics.IsMarkingSpeedIndex || productOutParametrics.IsMarkingCountry || productOutParametrics.IsMarkingTemperatureIndex ||
                        productOutParametrics.IsMarkingTractionIndex || productOutParametrics.IsMarkingTreadwearIndex || productOutParametrics.IsMarkingExtraLoad ||
                        productOutParametrics.IsMarkingFlangeProtection || productOutParametrics.IsMarkingRunFlat||productOutParametrics.IsMarkingAccomadation|| productOutParametrics.IsMarkingSpikes)
                    {
                        string[] arrStr = item.IdProduct.Split('-');
                        valuePairsDictionary = productOutParametrics.DictionarySrc.GetValuesById(arrStr[0], arrStr[1], arrStr[2]);
                    }


                    if (defaultOutParametrics.IsIdRows) CellSave(shareStringPart, row, cell, item.IdRow, cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsIdProduct) CellSave(shareStringPart, row, cell, item.IdProduct, cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsNameProduct) CellSave(shareStringPart, row, cell, item.NameProduct, cellList.NextVal() + index, shareIndex++);

                    if (productOutParametrics.IsBrandName) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_Name"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsBrandCountry) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_Country"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsBrandDescription) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_Description"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsBrandRunFlatName) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_RunFlatName"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelName) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Name"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelType) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Type"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelSeason) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Season"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelCommercial) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Commercial"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelDescription) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Description"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelWhileLetters) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_WhileLetters"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelMudSnow) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_MudSnow"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingWidth) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Width"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingHeight) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Height"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingDiameter) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Diameter"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingLoadIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_LoadIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingSpeedIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_SpeedIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingSpikes) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Spikes"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingCountry) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Country"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingTemperatureIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_TemperatureIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingTractionIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_TractionIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingTreadwearIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_TreadwearIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingExtraLoad) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_ExtraLoad"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingFlangeProtection) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_FlangeProtection"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingRunFlat) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_RunFlat"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingAccomadation) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Accomadation"], cellList.NextVal() + index, shareIndex++);
                    

                    if (defaultOutParametrics.IsIdPosition) CellSave(shareStringPart, row, cell, item.IdPosition, cellList.NextVal() + index, shareIndex++);
                    if (providerOutParametrics.IsProviderName) CellSave(shareStringPart, row, cell, (string)valuePairsProvider["Provider_Name"], cellList.NextVal() + index, shareIndex++);
                    if (providerOutParametrics.IsProviderPriority) CellSave(shareStringPart, row, cell, ((int)valuePairsProvider["Provider_Priority"]).ToString(), cellList.NextVal() + index, shareIndex++);
                    if (providerOutParametrics.IsStockName) CellSave(shareStringPart, row, cell, (string)valuePairsProvider["Stock_Name"], cellList.NextVal() + index, shareIndex++);
                    if (providerOutParametrics.IsStockTime) CellSave(shareStringPart, row, cell, (string)valuePairsProvider["Stock_Time"], cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsProviderPrice) CellSave(shareStringPart, row, cell, item.PriceProv.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsMarkup) CellSave(shareStringPart, row, cell, item.Markup.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsTotalPrice) CellSave(shareStringPart, row, cell, item.TotalPrice.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsCount) CellSave(shareStringPart, row, cell, item.Count.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsAddition) CellSave(shareStringPart, row, cell, item.Addition, cellList.NextVal() + index, shareIndex++);

                    sheetData.Append(row);

                    index++;
                }

                shareStringPart.SharedStringTable.Save();
                worksheet.Worksheet.Save();
            }
        }

        public void SaveDataExcel(string filepath, string[] idProviders)
        {
            SaveDataExcel(filepath, idProviders, new ExcelDefaultOutParametrics(), new ExcelProviderOutParametrics(), new ExcelProductOutParametrics());
        }

        public void SaveDataExcel(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics)
        {
            SaveDataExcel(filepath, idProviders, defaultOutParametrics, new ExcelProviderOutParametrics(), new ExcelProductOutParametrics());
        }

        public void SaveDataExcel(string filepath, string[] idProviders, ExcelProviderOutParametrics providerOutParametrics)
        {
            SaveDataExcel(filepath, idProviders, new ExcelDefaultOutParametrics(), providerOutParametrics, new ExcelProductOutParametrics());
        }

        public void SaveDataExcel(string filepath, string[] idProviders, ExcelProductOutParametrics productOutParametrics)
        {
            SaveDataExcel(filepath, idProviders, new ExcelDefaultOutParametrics(), new ExcelProviderOutParametrics(), productOutParametrics);
        }

        public void SaveDataExcel(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics, ExcelProviderOutParametrics providerOutParametrics)
        {
            SaveDataExcel(filepath, idProviders, defaultOutParametrics, providerOutParametrics, new ExcelProductOutParametrics());
        }

        public void SaveDataExcel(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics, ExcelProductOutParametrics productOutParametrics)
        {
            SaveDataExcel(filepath, idProviders, defaultOutParametrics, new ExcelProviderOutParametrics(), productOutParametrics);
        }

        public void SaveDataExcel(string filepath, string[] idProviders, ExcelProviderOutParametrics providerOutParametrics, ExcelProductOutParametrics productOutParametrics)
        {
            SaveDataExcel(filepath, idProviders, new ExcelDefaultOutParametrics(), providerOutParametrics, productOutParametrics);
        }





        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics, ExcelProviderOutParametrics providerOutParametrics, ExcelProductOutParametrics productOutParametrics)
        {
            List<DBSortedMarking> res = new List<DBSortedMarking>();
            foreach (var item in DBRows.Where(x => idProviders.Contains(x.IdPosition.Split('-').First())).OrderBy(x => x.IdProduct))
            {
                var itemRes = res.Find(x => x.IdProduct == item.IdProduct && x.IdPosition == item.IdPosition.Split('-').First() && x.Addition == item.Addition);

                if (itemRes != null)
                {
                    itemRes.Count += item.Count;
                    if (itemRes.TotalPrice < item.TotalPrice)
                    {
                        itemRes.TotalPrice = item.TotalPrice;
                        itemRes.PriceProv = item.PriceProv;
                    }
                }
                else
                {
                    res.Add(new DBSortedMarking()
                    {
                        IdProduct = item.IdProduct,
                        IdPosition = item.IdPosition.Split('-').First(),
                        NameProduct = item.NameProduct,
                        Markup = item.Markup,
                        PriceProv = item.PriceProv,
                        TotalPrice = item.TotalPrice,
                        Count = item.Count,
                        Addition = item.Addition
                        
                    });
                }
            }

            CreateExcel(filepath);

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true))
            {
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                if (shareStringPart.SharedStringTable == null)
                {
                    shareStringPart.SharedStringTable = new SharedStringTable();
                }

                int shareIndex = 0;

                WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                WorksheetPart worksheet;
                IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == "Data");
                if (sheets.Count() == 0)
                    throw new ArgumentException("Лист не найден");
                string relationshipId = sheets.First().Id.Value;
                worksheet = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheet.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                ExcelCellList cellList = new ExcelCellList();

                Cell cell = null;

                Row row = new Row() { RowIndex = 1 };

                if (defaultOutParametrics.IsIdProduct) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_IdProduct, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsNameProduct) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_NameProduct, cellList.NextVal() + "1", shareIndex++);

                if (productOutParametrics.IsBrandName) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandName, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsBrandCountry) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandCountry, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsBrandDescription) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandDescription, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsBrandRunFlatName) CellSave(shareStringPart, row, cell, productOutParametrics.Name_BrandRunFlatName, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelName) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelName, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelType) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelType, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelSeason) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelSeason, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelCommercial) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelCommercial, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelDescription) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelDescription, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelWhileLetters) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelWhileLetters, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsModelMudSnow) CellSave(shareStringPart, row, cell, productOutParametrics.Name_ModelMudSnow, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingWidth) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingWidth, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingHeight) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingHeight, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingDiameter) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingDiameter, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingLoadIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingLoadIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingSpeedIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingSpeedIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingSpikes) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingSpikes, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingCountry) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingCountry, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingTemperatureIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingTemperatureIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingTractionIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingTractionIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingTreadwearIndex) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingTreadwearIndex, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingExtraLoad) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingExtraLoad, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingFlangeProtection) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingFlangeProtection, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingRunFlat) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingRunFlat, cellList.NextVal() + "1", shareIndex++);
                if (productOutParametrics.IsMarkingAccomadation) CellSave(shareStringPart, row, cell, productOutParametrics.Name_MarkingAccomadation, cellList.NextVal() + "1", shareIndex++);
                

                if (defaultOutParametrics.IsIdPosition) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_IdPosition, cellList.NextVal() + "1", shareIndex++);
                if (providerOutParametrics.IsProviderName) CellSave(shareStringPart, row, cell, providerOutParametrics.Name_ProviderName, cellList.NextVal() + "1", shareIndex++);
                if (providerOutParametrics.IsProviderPriority) CellSave(shareStringPart, row, cell, providerOutParametrics.Name_ProviderPriority, cellList.NextVal() + "1", shareIndex++);

                if (defaultOutParametrics.IsProviderPrice) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_ProviderPrice, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsMarkup) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_Markup, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsTotalPrice) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_TotalPrice, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsCount) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_Count, cellList.NextVal() + "1", shareIndex++);
                if (defaultOutParametrics.IsAddition) CellSave(shareStringPart, row, cell, defaultOutParametrics.Name_Addition, cellList.NextVal() + "1", shareIndex++);

                sheetData.Append(row);

                Dictionary<string, object> valuePairsProvider = null;

                Dictionary<string, object> valuePairsDictionary = null;

                uint index = 2;
                foreach (var item in res)
                {
                    cellList.Restart();

                    row = new Row() { RowIndex = index };

                    if (providerOutParametrics.IsProviderName || providerOutParametrics.IsProviderPriority)
                    {
                        valuePairsProvider = providerOutParametrics.ProvidersSrc.GetValuesById(item.IdPosition);
                    }
                    if (productOutParametrics.IsBrandName || productOutParametrics.IsBrandCountry || productOutParametrics.IsBrandDescription || productOutParametrics.IsBrandRunFlatName ||
                        productOutParametrics.IsModelName || productOutParametrics.IsModelType || productOutParametrics.IsModelSeason || productOutParametrics.IsModelCommercial ||
                        productOutParametrics.IsModelDescription || productOutParametrics.IsModelWhileLetters || productOutParametrics.IsModelMudSnow ||
                        productOutParametrics.IsMarkingWidth || productOutParametrics.IsMarkingHeight || productOutParametrics.IsMarkingDiameter || productOutParametrics.IsMarkingLoadIndex ||
                        productOutParametrics.IsMarkingSpeedIndex || productOutParametrics.IsMarkingCountry || productOutParametrics.IsMarkingTemperatureIndex ||
                        productOutParametrics.IsMarkingTractionIndex || productOutParametrics.IsMarkingTreadwearIndex || productOutParametrics.IsMarkingExtraLoad ||
                        productOutParametrics.IsMarkingFlangeProtection || productOutParametrics.IsMarkingRunFlat||productOutParametrics.IsMarkingAccomadation|| productOutParametrics.IsMarkingSpikes)
                    {
                        string[] arrStr = item.IdProduct.Split('-');
                        valuePairsDictionary = productOutParametrics.DictionarySrc.GetValuesById(arrStr[0], arrStr[1], arrStr[2]);
                    }

                    if (defaultOutParametrics.IsIdProduct) CellSave(shareStringPart, row, cell, item.IdProduct, cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsNameProduct) CellSave(shareStringPart, row, cell, item.NameProduct, cellList.NextVal() + index, shareIndex++);

                    if (productOutParametrics.IsBrandName) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_Name"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsBrandCountry) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_Country"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsBrandDescription) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_Description"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsBrandRunFlatName) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Brand_RunFlatName"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelName) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Name"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelType) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Type"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelSeason) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Season"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelCommercial) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Commercial"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelDescription) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_Description"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelWhileLetters) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_WhileLetters"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsModelMudSnow) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Model_MudSnow"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingWidth) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Width"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingHeight) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Height"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingDiameter) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Diameter"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingLoadIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_LoadIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingSpeedIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_SpeedIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingSpikes) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Spikes"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingCountry) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Country"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingTemperatureIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_TemperatureIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingTractionIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_TractionIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingTreadwearIndex) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_TreadwearIndex"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingExtraLoad) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_ExtraLoad"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingFlangeProtection) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_FlangeProtection"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingRunFlat) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_RunFlat"], cellList.NextVal() + index, shareIndex++);
                    if (productOutParametrics.IsMarkingAccomadation) CellSave(shareStringPart, row, cell, (string)valuePairsDictionary["Marking_Accomadation"], cellList.NextVal() + index, shareIndex++);
                    

                    if (defaultOutParametrics.IsIdPosition) CellSave(shareStringPart, row, cell, item.IdPosition, cellList.NextVal() + index, shareIndex++);
                    if (providerOutParametrics.IsProviderName) CellSave(shareStringPart, row, cell, (string)valuePairsProvider["Provider_Name"], cellList.NextVal() + index, shareIndex++);
                    if (providerOutParametrics.IsProviderPriority) CellSave(shareStringPart, row, cell, ((int)valuePairsProvider["Provider_Priority"]).ToString(), cellList.NextVal() + index, shareIndex++);

                    if (defaultOutParametrics.IsProviderPrice) CellSave(shareStringPart, row, cell, item.PriceProv.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsMarkup) CellSave(shareStringPart, row, cell, item.Markup.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsTotalPrice) CellSave(shareStringPart, row, cell, item.TotalPrice.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsCount) CellSave(shareStringPart, row, cell, item.Count.ToString(), cellList.NextVal() + index, shareIndex++);
                    if (defaultOutParametrics.IsAddition) CellSave(shareStringPart, row, cell, item.Addition, cellList.NextVal() + index, shareIndex++);

                    sheetData.Append(row);

                    index++;
                }

                shareStringPart.SharedStringTable.Save();
                worksheet.Worksheet.Save();
            }
        }

        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders)
        {
            SaveDataOnlyProviderExcell(filepath, idProviders, new ExcelDefaultOutParametrics() { Name_IdPosition = "Id поставщик" }, new ExcelProviderOutParametrics(), new ExcelProductOutParametrics());
        }

        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics)
        {
            SaveDataOnlyProviderExcell(filepath, idProviders, defaultOutParametrics, new ExcelProviderOutParametrics(), new ExcelProductOutParametrics());
        }

        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders, ExcelProviderOutParametrics providerOutParametrics)
        {
            SaveDataOnlyProviderExcell(filepath, idProviders, new ExcelDefaultOutParametrics() { Name_IdPosition = "Id поставщик" }, providerOutParametrics, new ExcelProductOutParametrics());
        }

        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders, ExcelProductOutParametrics productOutParametrics)
        {
            SaveDataOnlyProviderExcell(filepath, idProviders, new ExcelDefaultOutParametrics() { Name_IdPosition = "Id поставщик" }, new ExcelProviderOutParametrics(), productOutParametrics);
        }

        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics, ExcelProviderOutParametrics providerOutParametrics)
        {
            SaveDataOnlyProviderExcell(filepath, idProviders, defaultOutParametrics, providerOutParametrics, new ExcelProductOutParametrics());
        }

        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders, ExcelDefaultOutParametrics defaultOutParametrics, ExcelProductOutParametrics productOutParametrics)
        {
            SaveDataOnlyProviderExcell(filepath, idProviders, defaultOutParametrics, new ExcelProviderOutParametrics(), productOutParametrics);
        }

        public void SaveDataOnlyProviderExcell(string filepath, string[] idProviders, ExcelProviderOutParametrics providerOutParametrics, ExcelProductOutParametrics productOutParametrics)
        {
            SaveDataOnlyProviderExcell(filepath, idProviders, new ExcelDefaultOutParametrics() { Name_IdPosition = "Id поставщик" }, providerOutParametrics, productOutParametrics);
        }
    }

    public class DBRow
    {
        public string IdRow { get; private set; }
        public string IdProduct { get; private set; }
        public string NameProduct { get; set; }
        public string IdPosition { get; private set; }

        public double PriceProv { get; set; }
        public double Markup { get; set; }
        public double TotalPrice { get; private set; }
        public int Count { get; set; }
        public string Addition;

        public DBRow(XmlNode x)
        {
            IdRow = x.Attributes.GetNamedItem("id").Value;
            IdProduct = x.SelectSingleNode("idProduct").InnerText;
            NameProduct = x.SelectSingleNode("nameProduct").InnerText;
            IdPosition = x.SelectSingleNode("idPosition").InnerText;
            PriceProv = double.Parse(x.SelectSingleNode("priceProvider").InnerText);
            Markup = double.Parse(x.SelectSingleNode("markup").InnerText);
            TotalPrice = double.Parse(x.SelectSingleNode("totalPrice").InnerText);
            Count = int.Parse(x.SelectSingleNode("count").InnerText);
            Addition = x.SelectSingleNode("addition").InnerText;
        }

        public DBRow(string idRow, string idProduct, string nameProduct, string idPosition, double priceProv, double markup,
             int count,string addition)
        {
            IdRow = idRow;
            IdProduct = idProduct;
            NameProduct = nameProduct;
            IdPosition = idPosition;
            PriceProv = priceProv;
            Markup = markup;
            TotalPrice = Math.Round((priceProv * (1 + markup / 100)) / 10, MidpointRounding.AwayFromZero) * 10;
            Count = count;
            Addition = addition;
        }


        internal XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("dRow");
            element.SetAttribute("id", IdRow);

            XmlElement e = xmlDocument.CreateElement("idProduct");
            e.InnerText = IdProduct;
            element.AppendChild(e);

            e = xmlDocument.CreateElement("nameProduct");
            e.InnerText = NameProduct;
            element.AppendChild(e);

            e = xmlDocument.CreateElement("idPosition");
            e.InnerText = IdPosition;
            element.AppendChild(e);

            e = xmlDocument.CreateElement("priceProvider");
            e.InnerText = PriceProv.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("markup");
            e.InnerText = Markup.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("totalPrice");
            e.InnerText = TotalPrice.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("count");
            e.InnerText = Count.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("addition");
            e.InnerText = Addition.ToString();
            element.AppendChild(e);

            return element;
        }
    }

    public class DBSortedMarking
    {
        public string IdProduct { get; set; }
        public string IdPosition { get; set; }
        public string NameProduct { get; set; }
        public double Markup { get; set; }
        public double PriceProv { get; set; }
        public double TotalPrice { get; set; }
        public int Count { get; set; }
        public string Addition { get; set; }
    }

    public class ExcelDefaultOutParametrics
    {
        public bool IsIdRows { get; set; }
        public bool IsIdProduct { get; set; }
        public bool IsNameProduct { get; set; }
        public bool IsIdPosition { get; set; }
        public bool IsProviderPrice { get; set; }
        public bool IsMarkup { get; set; }
        public bool IsTotalPrice { get; set; }
        public bool IsCount { get; set; }
        public bool IsAddition { get; set; }

        public string Name_IdRows { get; set; }
        public string Name_IdProduct { get; set; }
        public string Name_NameProduct { get; set; }
        public string Name_IdPosition { get; set; }
        public string Name_ProviderPrice { get; set; }
        public string Name_Markup { get; set; }
        public string Name_TotalPrice { get; set; }
        public string Name_Count { get; set; }
        public string Name_Addition { get; set; }

        public ExcelDefaultOutParametrics()
        {
            IsIdRows = false;
            IsIdProduct = true;
            IsNameProduct = true;
            IsIdPosition = true;
            IsProviderPrice = true;
            IsMarkup = true;
            IsTotalPrice = true;
            IsCount = true;
            IsAddition = true;

            SetNames();
        }

        public ExcelDefaultOutParametrics(bool isIdRows, bool isIdProduct, bool isNameProduct, bool isIdPosition, bool isProviderPrice,
            bool isMarkup, bool isTotalPrice, bool isCount, bool isAddition)
        {
            IsIdRows = isIdRows;
            IsIdProduct = isIdProduct;
            IsNameProduct = isNameProduct;
            IsIdPosition = isIdPosition;
            IsProviderPrice = isProviderPrice;
            IsMarkup = isMarkup;
            IsTotalPrice = isTotalPrice;
            IsCount = isCount;
            IsAddition = isAddition;

            SetNames();
        }

        private void SetNames()
        {
            SetNames("Id строки", "Id товара", "Наименование товара", "Id поставщик-склад", "Цена поставщика", "Наценка",
                "Итоговая цена", "Остаток","Примечание");
        }

        public void SetNames(string name_IdRows, string name_IdProduct, string name_NameProduct, string name_IdPosition, string name_ProviderPrice, string name_Markup,
            string name_TotalPrice, string name_Count,string name_Addition)
        {
            Name_IdRows = name_IdRows;
            Name_IdProduct = name_IdProduct;
            Name_NameProduct = name_NameProduct;
            Name_IdPosition = name_IdPosition;
            Name_ProviderPrice = name_ProviderPrice;
            Name_Markup = name_Markup;
            Name_TotalPrice = name_TotalPrice;
            Name_Count = name_Count;
            Name_Addition = name_Addition;
        }
    }

    public class ExcelProviderOutParametrics
    {
        public Providers ProvidersSrc { get; private set; }
        public bool IsProviderName { get; set; }
        public bool IsProviderPriority { get; set; }
        public bool IsStockName { get; set; }
        public bool IsStockTime { get; set; }

        public string Name_ProviderName { get; set; }
        public string Name_ProviderPriority { get; set; }
        public string Name_StockName { get; set; }
        public string Name_StockTime { get; set; }

        public ExcelProviderOutParametrics()
        {
            ProvidersSrc = null;
            IsProviderName = false;
            IsProviderPriority = false;
            IsStockName = false;
            IsStockTime = false;

            SetNames();
        }

        public ExcelProviderOutParametrics(Providers providersSrc, bool isProviderName, bool isProviderPriority, bool isStockName, bool isStockTime)
        {
            ProvidersSrc = providersSrc;
            IsProviderName = isProviderName;
            IsProviderPriority = isProviderPriority;
            IsStockName = isStockName;
            IsStockTime = isStockTime;

            SetNames();
        }

        private void SetNames()
        {
            SetNames("Название поставщика", "Приоритет поставщика", "Название склада", "Срок доставки");
        }

        public void SetNames(string name_ProviderName, string name_ProviderPriority, string name_StockName, string name_StockTime)
        {
            Name_ProviderName = name_ProviderName;
            Name_ProviderPriority = name_ProviderPriority;
            Name_StockName = name_StockName;
            Name_StockTime = name_StockTime;
        }
    }

    public class ExcelProductOutParametrics
    {
        public Dictionary DictionarySrc { get; set; }

        public bool IsBrandName { get; set; }
        public bool IsBrandCountry { get; set; }
        public bool IsBrandDescription { get; set; }
        public bool IsBrandRunFlatName { get; set; }
        public bool IsModelType { get; set; }
        public bool IsModelName { get; set; }
        public bool IsModelSeason { get; set; }
        public bool IsModelDescription { get; set; }
        public bool IsModelCommercial { get; set; }
        public bool IsModelWhileLetters { get; set; }
        public bool IsModelMudSnow { get; set; }
        public bool IsMarkingWidth { get; set; }
        public bool IsMarkingHeight { get; set; }
        public bool IsMarkingDiameter { get; set; }
        public bool IsMarkingSpeedIndex { get; set; }
        public bool IsMarkingLoadIndex { get; set; }
        public bool IsMarkingCountry { get; set; }
        public bool IsMarkingTractionIndex { get; set; }
        public bool IsMarkingTemperatureIndex { get; set; }
        public bool IsMarkingTreadwearIndex { get; set; }
        public bool IsMarkingExtraLoad { get; set; }
        public bool IsMarkingRunFlat { get; set; }
        public bool IsMarkingFlangeProtection { get; set; }
        public bool IsMarkingAccomadation { get; set; }
        public bool IsMarkingSpikes { get; set; }


        public string Name_BrandName { get; set; }
        public string Name_BrandCountry { get; set; }
        public string Name_BrandDescription { get; set; }
        public string Name_BrandRunFlatName { get; set; }
        public string Name_ModelType { get; set; }
        public string Name_ModelName { get; set; }
        public string Name_ModelSeason { get; set; }
        public string Name_ModelDescription { get; set; }
        public string Name_ModelCommercial { get; set; }
        public string Name_ModelWhileLetters { get; set; }
        public string Name_ModelMudSnow { get; set; }
        public string Name_MarkingWidth { get; set; }
        public string Name_MarkingHeight { get; set; }
        public string Name_MarkingDiameter { get; set; }
        public string Name_MarkingSpeedIndex { get; set; }
        public string Name_MarkingLoadIndex { get; set; }
        public string Name_MarkingCountry { get; set; }
        public string Name_MarkingTractionIndex { get; set; }
        public string Name_MarkingTemperatureIndex { get; set; }
        public string Name_MarkingTreadwearIndex { get; set; }
        public string Name_MarkingExtraLoad { get; set; }
        public string Name_MarkingRunFlat { get; set; }
        public string Name_MarkingFlangeProtection { get; set; }
        public string Name_MarkingAccomadation { get; set; }
        public string Name_MarkingSpikes { get; set; }

        public ExcelProductOutParametrics()
        {
            DictionarySrc = null;
            IsBrandName = false;
            IsBrandCountry = false;
            IsBrandDescription = false;
            IsBrandRunFlatName = false;
            IsModelType = false;
            IsModelName = false;
            IsModelSeason = false;
            IsModelDescription = false;
            IsModelCommercial = false;
            IsModelWhileLetters = false;
            IsModelMudSnow = false;
            IsMarkingWidth = false;
            IsMarkingHeight = false;
            IsMarkingDiameter = false;
            IsMarkingSpeedIndex = false;
            IsMarkingLoadIndex = false;
            IsMarkingCountry = false;
            IsMarkingTractionIndex = false;
            IsMarkingTemperatureIndex = false;
            IsMarkingTreadwearIndex = false;
            IsMarkingExtraLoad = false;
            IsMarkingRunFlat = false;
            IsMarkingFlangeProtection = false;
            IsMarkingSpikes = false;
            IsMarkingAccomadation = false;

            SetNames();
        }

        public ExcelProductOutParametrics(Dictionary dictionarySrc, bool isBrandName, bool isBrandCountry, bool isBrandDescription, bool isBrandRunFlatName, bool isModelType, bool isModelName,
            bool isModelSeason, bool isModelDescription, bool isModelCommercial, bool isModelWhileLetters, bool isModelMudSnow, bool isMarkingWidth, bool isMarkingHeight,
            bool isMarkingDiameter, bool isMarkingSpeedIndex, bool isMarkingLoadIndex, bool isMarkingCountry, bool isMarkingTractionIndex, bool isMarkingTemperatureIndex,
            bool isMarkingTreadwearIndex, bool isMarkingExtraLoad, bool isMarkingRunFlat, bool isMarkingFlangeProtection, bool isMarkingAccomadation, bool isMarkingSpikes)
        {
            DictionarySrc = dictionarySrc;
            IsBrandName = isBrandName;
            IsBrandCountry = isBrandCountry;
            IsBrandDescription = isBrandDescription;
            IsBrandRunFlatName = isBrandRunFlatName;
            IsModelType = isModelType;
            IsModelName = isModelName;
            IsModelSeason = isModelSeason;
            IsModelDescription = isModelDescription;
            IsModelCommercial = isModelCommercial;
            IsModelWhileLetters = isModelWhileLetters;
            IsModelMudSnow = isModelMudSnow;
            IsMarkingWidth = isMarkingWidth;
            IsMarkingHeight = isMarkingHeight;
            IsMarkingDiameter = isMarkingDiameter;
            IsMarkingSpeedIndex = isMarkingSpeedIndex;
            IsMarkingLoadIndex = isMarkingLoadIndex;
            IsMarkingCountry = isMarkingCountry;
            IsMarkingTractionIndex = isMarkingTractionIndex;
            IsMarkingTemperatureIndex = isMarkingTemperatureIndex;
            IsMarkingTreadwearIndex = isMarkingTreadwearIndex;
            IsMarkingExtraLoad = isMarkingExtraLoad;
            IsMarkingRunFlat = isMarkingRunFlat;
            IsMarkingFlangeProtection = isMarkingFlangeProtection;
            IsMarkingAccomadation = isMarkingAccomadation;
            IsMarkingSpikes = isMarkingSpikes;

            SetNames();
        }

        private void SetNames()
        {
            SetNames("Название производителя", "Страна производителя", "Описание производителя", "Название технологии RunFlat",
                "Тип модели", "Название модели", "Сезонность", "Описание модели", "Commercial", "БУКВЫ", "M+S", "Ширина", "Высота",
                "Диаметр", "Индекс скорости", "Индекс нагрузки", "Страна производства", "TractionIndex", "TemperatureIndex", "TreadwearIndex",
                "ExtraLoad", "Технология RunFlat", "FlangeProtection","Аккомадация","Шиповка");
        }

        public void SetNames(string name_BrandName, string name_BrandCountry, string name_BrandDescription, string name_BrandRunFlatName, string name_ModelType, string name_ModelName,
            string name_ModelSeason, string name_ModelDescription, string name_ModelCommercial, string name_ModelWhileLetters, string name_ModelMudSnow, string name_MarkingWidth,
            string name_MarkingHeight, string name_MarkingDiameter, string name_MarkingSpeedIndex, string name_MarkingLoadIndex, string name_MarkingCountry, string name_MarkingTractionIndex,
            string name_MarkingTemperatureIndex, string name_MarkingTreadwearIndex, string name_MarkingExtraLoad, string name_MarkingRunFlat, string name_MarkingFlangeProtection,
            string name_MarkingAccomadation, string name_MarkingSpikes)
        {
            Name_BrandName = name_BrandName;
            Name_BrandCountry = name_BrandCountry;
            Name_BrandDescription = name_BrandDescription;
            Name_BrandRunFlatName = name_BrandRunFlatName;
            Name_ModelType = name_ModelType;
            Name_ModelName = name_ModelName;
            Name_ModelSeason = name_ModelSeason;
            Name_ModelDescription = name_ModelDescription;
            Name_ModelCommercial = name_ModelCommercial;
            Name_ModelWhileLetters = name_ModelWhileLetters;
            Name_ModelMudSnow = name_ModelMudSnow;
            Name_MarkingWidth = name_MarkingWidth;
            Name_MarkingHeight = name_MarkingHeight;
            Name_MarkingDiameter = name_MarkingDiameter;
            Name_MarkingSpeedIndex = name_MarkingSpeedIndex;
            Name_MarkingLoadIndex = name_MarkingLoadIndex;
            Name_MarkingCountry = name_MarkingCountry;
            Name_MarkingTractionIndex = name_MarkingTractionIndex;
            Name_MarkingTemperatureIndex = name_MarkingTemperatureIndex;
            Name_MarkingTreadwearIndex = name_MarkingTreadwearIndex;
            Name_MarkingExtraLoad = name_MarkingExtraLoad;
            Name_MarkingRunFlat = name_MarkingRunFlat;
            Name_MarkingFlangeProtection = name_MarkingFlangeProtection;
            Name_MarkingAccomadation = name_MarkingAccomadation;
            Name_MarkingSpikes = name_MarkingSpikes;
        }
    }
}