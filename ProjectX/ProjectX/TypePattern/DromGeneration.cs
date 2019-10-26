using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ProjectX.ExcelParsing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace ProjectX.TypePattern
{
   public class DromGeneration
    {
        public static List<DromAd> Generate(List<Element> elements, GroupPattern groupPattern,
           int count, List<DromAd> dromAds)
        {

            Patterns patterns = new Patterns();
            Pattern pattern;

            count = elements.Count > count ? count : elements.Count;
            int curCount = 0;

            foreach (var item in elements)
            {

                pattern = patterns.GetPattern(groupPattern.GetIdPatter());

                string name = item.BrandName + " " + item.ModelName + " " + item.Accomadation;

                if (item.RunFlat == "Да") {
                    if (!(name.ToLower().Contains("runflat"))) {
                        name += " RunFlat";
                    }
                }

                string price = item.Price;

                if (int.Parse(item.Count) == 1)
                {
                    price = item.PriceForOne;
                }
                else if (int.Parse(item.Count) < 4)
                {
                    price = item.PriceForTwo;
                }


                var el = dromAds.Find(x => x.IdProduct == item.IdProduct);
                if (el == null)
                {
                    dromAds.Add(new DromAd()
                    {
                        IdProduct = item.IdProduct,
                        Diametr = item.Diameter,
                        Height = item.Height,
                        Seasson = item.Season,
                        Width = item.Width,
                        Priority = item.Priority,
                        Spikes = item.Spikes,
                        Head = GetLogicString(GetString(pattern.Head, item), item),
                        Body = GetLogicString(AddGetString(GetString(pattern.Body, item)), item),
                        Addition = item.Addition,
                        Price = price,
                        Images = item.Images,
                        IdProvider = item.ProvaiderId,
                        Brand = item.BrandName,
                        Time = item.TimeFromTo,
                        LoadIndex = item.LoadIndex,
                        SpeedIndex = item.SpeedIndex,
                        Count = item.Count,
                        Model = item.ModelName,
                        Commertial = item.Commercial,
                        Name = name

                    });

                    curCount++;
                }
                else if (el.Priority > item.Priority)
                {

                    dromAds.Remove(el);

                    dromAds.Add(new DromAd()
                    {
                        IdProduct = item.IdProduct,
                        Diametr = item.Diameter,
                        Height = item.Height,
                        Seasson = item.Season,
                        Width = item.Width,
                        Priority = item.Priority,
                        Spikes = item.Spikes,
                        Head = GetLogicString(GetString(pattern.Head, item), item),
                        Body = GetLogicString(AddGetString(GetString(pattern.Body, item)), item),
                        Addition = item.Addition,
                        Price = price,
                        Images = item.Images,
                        IdProvider = item.ProvaiderId,
                        Brand = item.BrandName,
                        Time = item.TimeFromTo,
                        LoadIndex = item.LoadIndex,
                        SpeedIndex = item.SpeedIndex,
                        Count = item.Count,
                        Model = item.ModelName,
                        Commertial = item.Commercial,
                        Name = name


                    });

                }

                if (curCount == count)
                {
                    break;
                }
            }

            return dromAds;
        }


        public static string GetString(string str, Element element)
        {

            str = str.Replace("<BrandDescription>", element.BrandDescription).Replace("<BDesc>", element.BrandDescription)
                .Replace("<Описание_производителя>", element.BrandDescription);
            str = str.Replace("<ModelDescription>", element.ModelDescription).Replace("<MDesc>", element.ModelDescription)
                .Replace("<Описание_модели>", element.ModelDescription);

            if (string.IsNullOrEmpty(element.Season))
            {
                str = str.Replace("<Season>", "").Replace("<LowSeason>", "").Replace("<UpSeason>", "")
                .Replace("<SeasonWihtSpikes>", "")
                .Replace("<LowSeasonWithSpikes>", "")
                .Replace("<UpSeasonWithSpikes>", "")
                .Replace("<Сезонность>", "").Replace("<Сезонность(мал.)>", "").Replace("<Сезоность(Бол.)>", "")
                .Replace("<Сезонность_с_шиповкой>", "")
                .Replace("<Сезонность_с_шиповкой(мал)>", "")
                .Replace("<Сезонность_с_шиповкой(Бол)>", "");
            }
            else if (element.Season != "")
            {
                str = str.Replace("<Season>", element.Season).Replace("<LowSeason>", element.Season.ToLower()).Replace("<UpSeason>", element.Season.First().ToString().ToUpper() + element.Season.Remove(0, 1))
                    .Replace("<Сезонность>", element.Season).Replace("<Сезонность(мал.)>", element.Season.ToLower()).Replace("<Сезоность(Бол.)>", element.Season.First().ToString().ToUpper() + element.Season.Remove(0, 1));
            }

            string com = element.Commercial == "Да" ? "C" : "";

            string seasWithSpikes = element.Season;

            string runFlat = element.ModelName.ToLower().Contains("runflat") ? "" : "RunFlat";

            if (seasWithSpikes == "Зимние")
            {
                if (element.Spikes == "Да")
                {
                    seasWithSpikes = "Зимние шипованные";
                }
                else
                {
                    seasWithSpikes = "Зимние нешипованные";
                }
            }


            str = str.Replace("<Accomadation>", element.Accomadation).Replace("<Ac>", element.Accomadation).Replace("<Амологация>", element.Accomadation)
                .Replace("<Addition>", element.Addition).Replace("<Add>", element.Addition).Replace("<Примичание>", element.Addition)
                .Replace("<BrandCountry>", element.BrandCountry).Replace("<BCountry>", element.BrandCountry).Replace("<Страна_производителя>", element.BrandCountry)
                .Replace("<BrandName>", element.BrandName).Replace("<BName>", element.BrandName).Replace("<Производитель>", element.BrandName)
                .Replace("<Commercial>", com).Replace("<C>", com).Replace("<Тип_шины(C)>", com)
                .Replace("<Count>", element.Count).Replace("<Ост>", element.Count).Replace("<Остаток>", element.Count)
                .Replace("<Date>", element.Date).Replace("<Дата>", element.Date)
                .Replace("<Diameter>", element.Diameter).Replace("<D>", element.Diameter).Replace("<d>", element.Diameter).Replace("<Диаметр>", element.Diameter).Replace("<Д>", element.Diameter).Replace("<д>", element.Diameter)
                .Replace("<ExtraLoad>", element.ExtraLoad.Replace("Да", "XL").Replace("Нет", "")).Replace("<XL>", element.ExtraLoad.Replace("Да", "XL").Replace("Нет", ""))
                .Replace("<xl>", element.ExtraLoad.Replace("Да", "xl").Replace("Нет", "")).Replace("<Повышенная_нагрузка>", element.ExtraLoad.Replace("Да", "XL").Replace("Нет", ""))
                .Replace("<FlangeProtection>", element.FlangeProtection).Replace("<FP>", element.FlangeProtection).Replace("<Защита борта>", element.FlangeProtection)
                .Replace("<Height>", element.Height).Replace("<H>", element.Height).Replace("<h>", element.Height).Replace("<Высота>", element.Height).Replace("<В>", element.Height).Replace("<в>", element.Height)
                .Replace("<LoadIndex>", element.LoadIndex).Replace("<LI>", element.LoadIndex).Replace("<Индекс_нагрузки>", element.LoadIndex)
                .Replace("<MarkingCountry>", element.MarkingCountry).Replace("<MCountry>", element.MarkingCountry).Replace("<Страна_производства>", element.MarkingCountry)
                .Replace("<ModelName>", element.ModelName).Replace("<MName>", element.ModelName).Replace("<Модель>", element.ModelName)
                .Replace("<MudSnow>", element.MudSnow.Replace("Да", "M+S").Replace("Нет", "")).Replace("<M+S>", element.MudSnow.Replace("Да", "M+S").Replace("Нет", ""))
                .Replace("<Price>", element.Price).Replace("<$>", element.Price).Replace("<Цена>", element.Price)
                .Replace("<RunFlat>", element.RunFlat.Replace("Да", runFlat).Replace("Нет", "")).Replace("<RF>", element.RunFlat.Replace("Да", runFlat).Replace("Нет", ""))
                .Replace("<RunFlatName>", element.RunFlatName).Replace("<Название_технологии_RunFlat>", element.RunFlatName)
                .Replace("<SeasonWihtSpikes>", seasWithSpikes).Replace("<Сезонность_с_шиповкой>", seasWithSpikes)
                .Replace("<LowSeasonWithSpikes>", seasWithSpikes.ToLower()).Replace("<Сезонность_с_шиповкой(мал)>", seasWithSpikes.ToLower())
                .Replace("<UpSeasonWithSpikes>", seasWithSpikes.First().ToString().ToUpper() + seasWithSpikes.Remove(0, 1)).Replace("<Сезонность_с_шиповкой(Бол)>", seasWithSpikes.First().ToString().ToUpper() + seasWithSpikes.Remove(0, 1))

                .Replace("<SpeedIndex>", element.SpeedIndex).Replace("<SI>", element.SpeedIndex).Replace("<Индекс_скорости>", element.SpeedIndex)
                .Replace("<Spikes>", element.Spikes.Replace("Да", "шипованные").Replace("Нет", "нешипованные"))
                .Replace("<Шип>", element.Spikes.Replace("Да", "шипованные").Replace("Нет", "нешипованные"))
                .Replace("<UpSpikes>", element.Spikes.Replace("Да", "Шипованные").Replace("Нет", "Нешипованные"))
                .Replace("<Шип(Бол)>", element.Spikes.Replace("Да", "Шипованные").Replace("Нет", "Нешипованные"))


                .Replace("<TemperatureIndex>", element.TemperatureIndex).Replace("<Temperature>", element.TemperatureIndex).Replace("<Индекс_температуры>", element.TemperatureIndex)
                .Replace("<Time>", element.Time).Replace("<Время>", element.Time)
                .Replace("<TimeTransit>", element.TimeTransit).Replace("<Срок_доставки>", element.TimeTransit)
                .Replace("<TractionIndex>", element.TractionIndex).Replace("<Traction>", element.TractionIndex).Replace("<Индекс_сцепления>", element.TractionIndex)
                .Replace("<TreadwearIndex>", element.TreadwearIndex).Replace("<Treadwear>", element.TreadwearIndex).Replace("<Индекс_ходимости>", element.TreadwearIndex)
                .Replace("<Type>", element.Type).Replace("<Тип_шины>", element.Type)
                .Replace("<WhileLetters>", element.WhileLetters).Replace("<WL>", element.WhileLetters).Replace("<Рефленные буквы>", element.WhileLetters)
                .Replace("<Width>", element.Width).Replace("<W>", element.Width).Replace("<w>", element.Width).Replace("<Ширина>", element.Width).Replace("<ш>", element.Width).Replace("<Ш>", element.Width)
                ;




            return str;
        }

        public static string AddGetString(string str)
        {
            str = str.Replace("<Жирный>", "").Replace("<Ж>", "").Replace("</Жирный>", "").Replace("</Ж>", "").Replace("<strong>", "").Replace("</strong>", "")
                .Replace("<Курсив>", "").Replace("<К>", "").Replace("</Курсив>", "").Replace("</К>", "").Replace("<K>", "").Replace("</K>", "").Replace("<em>", "").Replace("</em>", "")
                .Replace("<Маркированный список>", "").Replace("<MC>", "").Replace("<ul>", "").Replace("</Маркированный список>", "").Replace("</MC>", "").Replace("</ul>", "")
                .Replace("<Нумерованный список>", "").Replace("<НC>", "").Replace("<ol>", "").Replace("</Нумерованный список>", "").Replace("</НC>", "").Replace("</ol>", "")
                .Replace("<Элемемнт списка>", "").Replace("<ЭС>", "").Replace("<li>", "").Replace("</Элемемнт списка>", "").Replace("</ЭС>", "").Replace("</li>", "")
                .Replace("<p>", "").Replace("</p>", "").Replace("<br>", "");
            ;



            return str;

        }

        public static string GetLogicString(string str, Element element)
        {
            MatchCollection matchCollection = Regex.Matches(str, "\\{\\[[^\\[\\]\\{\\}\\(\\)]*\\][!=]='[^\\[\\]\\{\\}\\(\\)]*'\\([^\\[\\]\\{\\}\\(\\)]*\\)(\\([^\\[\\]\\{\\}\\(\\)]*\\))?\\}");
            foreach (Match item in matchCollection)
            {
                string val = item.Value;

                string param = Regex.Match(val, "\\[[^\\[\\]\\{\\}\\(\\)]*\\]").Value;
                param = param.Remove(param.Length - 1, 1).Remove(0, 1);

                string logic = Regex.Match(val, "[!=]=").Value;

                string lparam = Regex.Match(val, "'[^\\[\\]\\{\\}\\(\\)]*'").Value;
                lparam = lparam.Remove(lparam.Length - 1, 1).Remove(0, 1);

                string value1 = Regex.Matches(val, "\\([^\\[\\]\\{\\}\\(\\)]*\\)")[0].Value;
                value1 = value1.Remove(value1.Length - 1, 1).Remove(0, 1);
                string value2 = "";

                if (Regex.Matches(val, "\\([^\\[\\]\\{\\}\\(\\)]*\\)").Count > 1)
                {
                    value2 = Regex.Matches(val, "\\([^\\[\\]\\{\\}\\(\\)]*\\)")[1].Value;
                    value2 = value2.Remove(value2.Length - 1, 1).Remove(0, 1);
                }



                string valparam = element.GetParam(param);

                lparam = lparam.Trim() == "@NULL" ? "" : lparam;

                if (logic == "==")
                {
                    if (valparam == lparam)
                    {
                        str = str.Replace(val, value1);
                    }
                    else
                    {
                        str = str.Replace(val, value2);
                    }
                }
                else if (logic == "!=")
                {
                    if (valparam != lparam)
                    {
                        str = str.Replace(val, value1);
                    }
                    else
                    {
                        str = str.Replace(val, value2);
                    }
                }
                else
                {
                    str = str.Replace(val, "");
                }




            }

            return str;
        }

        private static void CreateExcel(string filepath)
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
                Name = "Лист1"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            spreadsheetDocument.Close();
        }

        private static void CellSave(SharedStringTablePart shareStringPart, Row row, Cell cell, string value, string cellReference, int index)
        {
            shareStringPart.SharedStringTable.AppendChild(
                         new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(value)));
            cell = new Cell() { CellReference = cellReference };
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            row.AppendChild(cell);
        }

        private static void CellSave(SharedStringTablePart shareStringPart, Row row, Cell cell, string value, string cellReference)
        {
            CellSave(shareStringPart, row, cell, value, cellReference, shareStringPart.SharedStringTable.Count());
        }


        public static void ToXLSX(string path, List<DromAd> dromAds)
        {

            CreateExcel(path);

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(path, true))
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
                IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == "Лист1");
                if (sheets.Count() == 0)
                    throw new ArgumentException("Лист не найден");
                string relationshipId = sheets.First().Id.Value;
                worksheet = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheet.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                ExcelCellList cellList = new ExcelCellList();

                Cell cell = null;

                Row row = new Row() { RowIndex = 1 };

                CellSave(shareStringPart, row, cell, "Артикул", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Товар", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Наименование", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Индекс нагрузки", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Индекс скорости", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Маркировка", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Сезонность", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Шиповка", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Тип шины", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Тип диска", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Остаток", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Продажа", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Описание", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Цена", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Состояние", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Наличие", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Срок доставки", cellList.NextVal() + "1", shareIndex++);
                CellSave(shareStringPart, row, cell, "Картинка", cellList.NextVal() + "1", shareIndex++);

                sheetData.Append(row);


                uint index = 2;
                foreach (var item in dromAds)
                {

                    string seasson = "";
                    string spikes = "";
                    string type = item.Commertial == "Легкогрузовая" ? "": "Обычная";

                    string buf = "";

                    if (Convert.ToInt32(item.Count) < 4)
                    {
                        buf = item.Count;
                    }
                    else if (Convert.ToInt32(item.Count) < 6)
                    {
                        buf = "4";
                    }
                    else if (Convert.ToInt32(item.Count) < 8)
                    {
                        buf = "6";
                    }
                    else
                    {
                        buf = "8";
                    }



                    if (item.Seasson == "Летние")
                    {
                        seasson = "Летняя";
                    }
                    else if (item.Seasson == "Зимние")
                    {
                        seasson = "Зимняя";
                        if (item.Spikes == "Да") {
                            spikes = "шипованная";
                        }
                    }
                    else if (item.Seasson == "Всесезонные")
                    {
                        seasson = "Всесезонная";
                    }
                    else {
                        continue;
                    }

                    cellList.Restart();


                    row = new Row() { RowIndex = index };

                    string image = "";

                    if (item.Images.Count > 0) {
                        image = item.Images.First();
                    }

                    var s = item.Addition == "" ? "" : item.Addition.GetHashCode().ToString();

                    

                    CellSave(shareStringPart, row, cell, item.IdProduct.GetHashCode().ToString() + s, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, "Шина", cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.Name, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.LoadIndex, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.SpeedIndex, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.Width+"/"+item.Height + " R"+item.Diametr, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, seasson, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, spikes, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, type, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, "", cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, buf, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, "шт", cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.Body, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.Price, cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, "новая", cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.Time == 0 ? "В наличии" : "Под заказ", cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, item.Time == 0? "": item.Time.ToString(), cellList.NextVal() + index, shareIndex++);
                    CellSave(shareStringPart, row, cell, image, cellList.NextVal() + index, shareIndex++);

                    sheetData.Append(row);

                    index++;
                }

            }

        }
    }


    public class DromAd
    {

        public string IdProduct;
        public string IdProvider;
        public int Priority;
        public string Addition;
        public string Commertial;
        public string Name;

        public string Width;
        public string Height;
        public string Diametr;
        public string LoadIndex;
        public string SpeedIndex;
        public string Spikes;
        public string Seasson;
        public string Price;
        public string Count;

        public string Brand;
        public string Model;

        public List<string> Images;

        public string Head { get; set; }
        public string Body;

        public int Time;
    }

}
