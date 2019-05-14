using ProjectX.AnalysisType;
using ProjectX.DataBase;
using ProjectX.Dict;
using ProjectX.ExcelParsing;
using ProjectX.Information;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ProjectX.TypePattern;


namespace ProjectX
{
    internal class Program
    {
        public static void Main(string[] args)
        {
           // Test1();



            Test3();
        }

        private static void Test1()
        {
            Console.WriteLine("********************************************************************\n");
            Console.WriteLine("-------------------------TEST1--------------------------------------\n");

            var fileName1 = @"C:\Users\ACER\Desktop\Прайсы\ВячеСлавик.xlsx";
            var fileName2 = @"C:\Users\ACER\Desktop\Прайсы\4точки.xlsx";

            if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data")))
            {
                Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data"));
            }

            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Dictionary.xml")))
            {
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Dictionary.xml"));
            }
            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\DictLog.txt")))
            {
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\DictLog.txt"));
            }
            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\ProvidersRegulars.xml")))
            {
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\ProvidersRegulars.xml"));
            }
            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Providers.xml")))
            {
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Providers.xml"));
            }
            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\DataBase.xml")))
            {
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\DataBase.xml"));
            }

            Stopwatch sw = new Stopwatch();
            sw.Start();

            Providers providers = new Providers();

            sw.Stop();
            Console.WriteLine("Инициализация поставщиков XML (пустой)" + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            string id = providers.AddProvider("ВячеСлавик", 1);

            sw.Stop();
            Console.WriteLine("Добавили поставщика " + id + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            string ids = providers.AddStock(id, "Основная группа", "1дн.");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Ставрополь", "1дн.");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Склады 2 дня", "3дн.");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Склады 7 дней", "1нед.1дн.");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            id = providers.AddProvider("4точки", 2);
            sw.Stop();
            Console.WriteLine("Добавили поставщика " + id + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Краснодар", "1дн.");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Москва", "1нед.");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            ids = providers.AddStock(id, "Домодедово", "1нед.");
            ids = providers.AddStock(id, "Москва склад 2", "1нед.");
            ids = providers.AddStock(id, "Москва склад 3", "1нед.");
            ids = providers.AddStock(id, "Москва склад 4", "1нед.");

            sw.Restart();
            providers.Save();
            sw.Stop();
            Console.WriteLine("Сохранение поставщиков в XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            providers = new Providers();
            sw.Stop();
            Console.WriteLine("Инициализация поставщиков XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            List<EParsingParam> eParsingParams = new List<EParsingParam>();
            string[] bufIndex1 = { "B" };
            
            EParsingParam parsingParam1 = new EParsingParam(fileName1, "A0");
            parsingParam1.AddStringVal("более 20", 20);
            parsingParam1.AddStringVal("более 40", 40);
            ESheet eSheet = new ESheet(null, "D");
            
            eSheet.AddBufIndex(bufIndex1);
            eSheet.AddCountIndex("A0", "E");
            eSheet.AddCountIndex("A1", "F");
            eSheet.AddCountIndex("A2", "G");
            eSheet.AddCountIndex("A3", "H");
            parsingParam1.Add(eSheet);

            eParsingParams.Add(parsingParam1);
            

            
            parsingParam1 = new EParsingParam(fileName2, id);
            parsingParam1.AddStringVal("более 20", 20);
            parsingParam1.AddStringVal("более 40", 40);

            string[] bufIndex2 = { "B", "D","J","K","L","M"};
            EReplacementCell repCell1 = new EReplacementCell("J");
            repCell1.Add("Да", "Шипов");
            EReplacementCell repCell2 = new EReplacementCell("K");
            repCell2.Add("Да", "XL");
            EReplacementCell repCell3 = new EReplacementCell("M");
            repCell3.Add("Да", "RUNFLAT");
            EReplacementCell repCell4 = new EReplacementCell("L");
            repCell4.Add("M+S 3PMSF", "M+S");
            repCell4.Add("#", "");
            EReplacementCell repCell5 = new EReplacementCell("D");
            repCell5.Add("Nordman", "Nokian");

            eSheet = new ESheet("1 Шины (Краснодар Индустриальны", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddReplacementCell(repCell1);
            eSheet.AddReplacementCell(repCell2);
            eSheet.AddReplacementCell(repCell3);
            eSheet.AddReplacementCell(repCell4);
            eSheet.AddReplacementCell(repCell5);
            eSheet.AddCountIndex("A0", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("3 Шины (Москва)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddReplacementCell(repCell1);
            eSheet.AddReplacementCell(repCell2);
            eSheet.AddReplacementCell(repCell3);
            eSheet.AddReplacementCell(repCell4);
            eSheet.AddReplacementCell(repCell5);
            eSheet.AddCountIndex("A1", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("4 Шины (Домодедово)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddReplacementCell(repCell1);
            eSheet.AddReplacementCell(repCell2);
            eSheet.AddReplacementCell(repCell3);
            eSheet.AddReplacementCell(repCell4);
            eSheet.AddReplacementCell(repCell5);
            eSheet.AddCountIndex("A2", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("5 Шины (Склад 2)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddReplacementCell(repCell1);
            eSheet.AddReplacementCell(repCell2);
            eSheet.AddReplacementCell(repCell3);
            eSheet.AddReplacementCell(repCell4);
            eSheet.AddReplacementCell(repCell5);
            eSheet.AddCountIndex("A3", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("6 Шины (Склад 3)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddReplacementCell(repCell1);
            eSheet.AddReplacementCell(repCell2);
            eSheet.AddReplacementCell(repCell3);
            eSheet.AddReplacementCell(repCell4);
            eSheet.AddReplacementCell(repCell5);
            eSheet.AddCountIndex("A4", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("7 Шины (Склад 4)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddReplacementCell(repCell1);
            eSheet.AddReplacementCell(repCell2);
            eSheet.AddReplacementCell(repCell3);
            eSheet.AddReplacementCell(repCell4);
            eSheet.AddReplacementCell(repCell5);
            eSheet.AddCountIndex("A5", "S");
            parsingParam1.Add(eSheet);
            
            eParsingParams.Add(parsingParam1);
            

            ProviderRegulars providerRegulars = new ProviderRegulars();
            providerRegulars.Add("A0");

            string idR = providerRegulars.Add("A0", "[0-9]{1,3}([.,][0-9]{1,3})?[\\*/][0-9]{1,3}([.,][0-9]{1,3})?R[0-9]{2}[CС]?", 0);

            ProviderRegular PR = providerRegulars.GetProviderRegularById("A0");
            PrimaryRegular PR2 = PR.GetPrimaryRegularById(idR);
            PR2.Add("C", 0, "commercial");
            PR2.Add("[0-9]{1,3}([.,][0-9]{1,3})?", 0, "width");
            PR2.Add("[0-9]{1,3}([.,][0-9]{1,3})?", 1, "height");
            string helpid = PR2.AddGroupRegular("R[0-9]{2}");
            GroupRegular GR = PR2.GetGroupRegularById(helpid);
            GR.Add("[0-9]{2}", 0, "diameter");

            idR = providerRegulars.Add("A0", "[0-9]{1,3}([.,][0-9]{1,3})?R[0-9]{2}[CС]?", 0);
            PR = providerRegulars.GetProviderRegularById("A0");
            PR2 = PR.GetPrimaryRegularById(idR);
            PR2.Add("C", 0, "commercial");
            PR2.Add("Полно профильные", "height");
            PR2.Add("[0-9]{1,3}([.,][0-9]{1,3})?", 0, "width");
            helpid = PR2.AddGroupRegular("R[0-9]{2}");
            GR = PR2.GetGroupRegularById(helpid);
            GR.Add("[0-9]{2}", 0, "diameter");


            idR = providerRegulars.Add("A0", "[0-9]{2,3}(/[0-9]{2,3})?[A-Z]{1,2}", 1);
            providerRegulars.Add("A0", idR, "[0-9]{2,3}(/[0-9]{2,3})?", 0, "loadIndex");
            providerRegulars.Add("A0", idR, "[A-Z]{1,2}", 0, "speedIndex");

            idR = providerRegulars.Add("A0", "XL", 2);
            providerRegulars.Add("A0", idR, "XL", 0, "extraLoad");

            idR = providerRegulars.Add("A0", "Шип", 3);
            providerRegulars.Add("A0", idR, "Да", "spikes");


            idR = providerRegulars.Add("A0", "\\*\\*\\*", 4);
            providerRegulars.Add("A0", idR, "Более 3-х лет", "additional");

            idR = providerRegulars.Add("A0", "\\([0-9]{2,4}г?\\)", 4);
            providerRegulars.Add("A0", idR, "\\([0-9]{2,4}г?\\)", 0, "additional");

            idR = providerRegulars.Add("A0", "\\*", 5);
            providerRegulars.Add("A0", idR, "BMW", "accomadation");


            idR = providerRegulars.Add("A0", "✩", 5);
            providerRegulars.Add("A0", idR, "BMW", "accomadation");


            idR = providerRegulars.Add("A0", "ROF RUN FLAT", 6);
            providerRegulars.Add("A0", idR, "Yes", "runFlat");

            idR = providerRegulars.Add("A0", "RUN FLAT", 6);
            providerRegulars.Add("A0", idR, "Yes", "runFlat");






            providerRegulars.Add("A1");

            idR = providerRegulars.Add("A1", "(LT)?[0-9]{1,3}([.,][0-9]{1,3})?[/x][0-9]{1,3}([.,][0-9]{1,3})?Z?R[0-9]{2}[CС]?", 0);

            PR = providerRegulars.GetProviderRegularById("A1");
            PR2 = PR.GetPrimaryRegularById(idR);
            PR2.Add("C", 0, "commercial");
            PR2.Add("ZR", 0, "flangeProtection");
            PR2.Add("LT", 0, "commercial");
            PR2.Add("[0-9]{1,3}([.,][0-9]{1,3})?", 0, "width");
            PR2.Add("[0-9]{1,3}([.,][0-9]{1,3})?", 1, "height");
            helpid = PR2.AddGroupRegular("R[0-9]{2}");
            GR = PR2.GetGroupRegularById(helpid);
            GR.Add("[0-9]{2}", 0, "diameter");

            idR = providerRegulars.Add("A1", "(LT)?[0-9]{1,3}([.,][0-9]{1,3})?(Z)?R[0-9]{2}[CС]?", 0);
            PR = providerRegulars.GetProviderRegularById("A1");
            PR2 = PR.GetPrimaryRegularById(idR);
            PR2.Add("C", 0, "commercial");
            PR2.Add("ZR", 0, "flangeProtection");
            PR2.Add("LT", 0, "commercial");
            PR2.Add("Полно профильные", "height");
            PR2.Add("[0-9]{1,3}([.,][0-9]{1,3})?", 0, "width");
            helpid = PR2.AddGroupRegular("R[0-9]{2}");
            GR = PR2.GetGroupRegularById(helpid);
            GR.Add("[0-9]{2}", 0, "diameter");


            idR = providerRegulars.Add("A1", "[0-9]{2,3}(/[0-9]{2,3})?[A-Z]{1,2}", 1);
            providerRegulars.Add("A1", idR, "[0-9]{2,3}(/[0-9]{2,3})?", 0, "loadIndex");
            providerRegulars.Add("A1", idR, "[A-Z]{1,2}", 0, "speedIndex");

            idR = providerRegulars.Add("A1", "XL", 2);
            providerRegulars.Add("A1", idR, "XL", 0, "extraLoad");

            idR = providerRegulars.Add("A1", "Шипов", 3);
            providerRegulars.Add("A1", idR, "Шипов", "spikes");

            idR = providerRegulars.Add("A1", "RUNFLAT", 4);
            providerRegulars.Add("A1", idR, "RUNFLAT", 0, "runFlat");

            idR = providerRegulars.Add("A1", "M\\+S", 5);
            providerRegulars.Add("A1", idR, "M\\+S", 0, "mudSnow");

            providerRegulars.AddPassString("A1", "TL");
            providerRegulars.AddPassString("A1", "(шип.)");


            providerRegulars.Save();



            sw.Stop();
            Console.WriteLine("Подготовили параметры парсинга " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();

            List<ParsingRow> parsings = EParsing.GetParsingRows(eParsingParams.ToArray(), out List<string> warning);

            sw.Stop();
            Console.WriteLine("Парсинг файлов xlsx " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            Dictionary dictionary = new Dictionary();
            sw.Stop();
            Console.WriteLine("Инициализация словаря XML (пустой) " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dictionary.GetItemFromTxt(@"C:\Users\ACER\Desktop\Прайсы\Dictionary.txt");
            sw.Stop();
            Console.WriteLine("Подгрузка словаря из файла TXT " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dictionary.Save();

            sw.Stop();
            Console.WriteLine("Сохранение словаря в XML " + sw.ElapsedMilliseconds + " мс");

            dictionary.Close();

            sw.Restart();
            dictionary = new Dictionary();
            sw.Stop();
            Console.WriteLine("Инициализация словаря XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            var res = dictionary.Analysis(ref parsings,true);
            ParsingRow[] rows = res.Where(x => x.Resault is GResault).ToArray();
            ParsingRow[] bRows = res.Where(x => x.Resault is BResault).ToArray();
            ParsingRow[] NRows = res.Where(x => x.Resault is NResault).ToArray();
            sw.Stop();
            Console.WriteLine("Анализ данных " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            DBase dBase = new DBase();
            sw.Stop();
            Console.WriteLine("Инициализация базы XML(пустой) " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.AddRows(rows);
            sw.Stop();
            Console.WriteLine("Добавление в базу(оперативка) " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.Save();
            sw.Stop();
            Console.WriteLine("Сохранение базы XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase = new DBase();
            sw.Stop();
            Console.WriteLine("Инициализация базы XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.SaveDataOnlyUniqueProviderExcell(@"C:\Users\ACER\Desktop\Прайсы\Data.xlsx", new string[] { "A1","A0" }, new ExcelDefaultOutParametrics() {
                IsNameProduct = false

            },new ExcelProviderOutParametrics() {ProvidersSrc = providers,
                IsStockTime = true,
                IsStockName = true } ,
                new ExcelProductOutParametrics() { DictionarySrc = dictionary,
                    IsMarkingHeight = true,
                    IsMarkingDiameter = true,
                    IsMarkingWidth = true,
                    IsMarkingSpeedIndex = true,
                    IsMarkingExtraLoad = true,
                    IsMarkingLoadIndex = true,
                    IsMarkingAccomadation = true,
                    IsMarkingRunFlat = true,
                    IsBrandName = true,
                    IsModelName = true,
                    IsMarkingSpikes = true

                });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data.xlsx' " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();

            dBase.SaveDataOnlyProviderExcell(@"C:\Users\ACER\Desktop\Прайсы\Data3.xlsx", new string[] { "A0", "A1" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data3.xlsx' " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            EParsingParam parsingParam2 = new EParsingParam(fileName1, "A0");
            ESheet eSheet2 = new ESheet(null, "D");
            eSheet2.AddBufIndex(bufIndex1);
            eSheet2.AddCountIndex("A1", "F");
            parsingParam2.Add(eSheet2);
            sw.Stop();

            dictionary.Save();

            dictionary.Close();

            Console.WriteLine("готово");
            Console.WriteLine("********************************************************************\n\n\n");
            Console.ReadKey();
        }


        private static void Test3() {

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data")))
            {
                Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data"));
            }

            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Patterns.xml")))
            {
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Patterns.xml"));
            }

            Patterns patterns = new Patterns();

            patterns.Add("<Up_Season> шины R<d>{[C]=='Да'(<C>)} <w>{[h]!='Полно профильные'( <h>)( )} <BName> <MName>",
                "<strong><Up_Season> шины R<d>{[C]=='Да'(<C>)} <w>{[h]!='Полно профильные'( <h>)( )} <BName> <MName> <LI><SI> <XL></strong>" + Environment.NewLine
                + "Новые <Low_Season>{[Spikes]=='Да'(шипованные)} <w>{[h]!='Полно профильные'(/<h>)( )} R <d>." + Environment.NewLine 
                + "Заявленную <strong>ЦЕНУ ГАРАНТИРУЕМ</strong> при покупке полного комплекта из четырех колес при оплате за наличный расчет!" + Environment.NewLine + Environment.NewLine
                + "Нашли дешевле - сообщите нам номер объявления и мы предложим Вам лучшую цену!" + Environment.NewLine + Environment.NewLine
                + "Наличие: в наличии на складе на <Date> <Time> мск. – звоните!" + Environment.NewLine
                + "Заказ и <strong>ДОСТАВКА</strong> курьером <strong>ПО КРАСНОДАРУ И КРАСНОДАРСКОМУ КРАЮ</strong> без предоплаты!" + Environment.NewLine + Environment.NewLine
                + "Работаем с юр.лицами с <strong>НДС!</strong>" + Environment.NewLine
                + "Новые <Low_Season> шины <w>{[h]!='Полно профильные'( <h>)( )} <d>{[C]=='Да'(<C>)} eсть в наличии других производителей." + Environment.NewLine + Environment.NewLine
                + "<Up_Season> шины <w>{[h]!='Полно профильные'(/<h>)( )}R<d> БУ не продаём!" + Environment.NewLine + Environment.NewLine
                + "Указана цена за 1 штуку." + Environment.NewLine 
                + "Расшифровка индекса скорости " + Environment.NewLine 
                + "<ul><li>T - 225-235 км/ч</li><li>U - 200-205 км/ч</li><li>H - 210-215 км/ч</li><li>V - 240-245 км/ч</li><li>W - 255-265 км/ч</li>" +
                "<li>Y - 275-285 до 295 км/ч</li><li>Z - 305-315 км/ч до 325-335</li></ul>");

            patterns.Save();

            AvitoConfig config = new AvitoConfig();
            config.Add("летние", false, "Летние");
            config.Add("зимние", true, "Зимние шипованные");
            config.Add("зимние", false, "Зимние нешипованные");
            config.Add("всесезонные", false, "Всесезонные");
            config.Save();

            patterns = new Patterns();

            foreach (Pattern item in patterns)
            {
              //  Console.WriteLine("ID: " + item.Id);
              //  Console.WriteLine("Загаловок: " + item.Head);
              //  Console.WriteLine("Тело: " + item.Body);
            }

            List<Element> elements1 = Element.GetElements();

            Providers providers = new Providers();
            Dictionary dictionary = new Dictionary();


            Filter filter = new Filter()
            {
                ProvidersId = providers.GetId(),
               BrandsId = dictionary.GetBrandsId(),
              // BrandsId = { "B2" },
               Widths = dictionary.GetWidths(),
               //Widths = {"185"},
               Heights = dictionary.GetHeights(),
               Diameters = dictionary.GetDiametrs(),
               Spikes = { "Да", "Нет" },
               Seassons = dictionary.GetSeassons()

                
            };

            dictionary.Close();

            List<Element>  elements = filter.GetElements(elements1);

            elements = Element.Distinct(elements);

            List<AvitoAd> avitoAds = AvitoGenerator.Generate(elements, "A0", 40, new List<AvitoAd>());

            /*
            foreach (var item in avitoAds)
            {
                Console.WriteLine(item.Head);
                Console.WriteLine();
                Console.WriteLine(item.Body);
            }
            */

            AvitoGenerator.ToXML(@"C:\Users\ACER\Desktop\TEST_AVITO.xml", avitoAds);

            stopwatch.Stop();

            Console.WriteLine(stopwatch.ElapsedMilliseconds.ToString());
            Console.ReadLine();
        }
    }
}