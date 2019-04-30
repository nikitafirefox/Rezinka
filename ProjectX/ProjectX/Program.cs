using ProjectX.DataBase;
using ProjectX.Dict;
using ProjectX.ExcelParsing;
using ProjectX.Information;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ProjectX
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            Test1();

           // Test2();
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
            string ids = providers.AddStock(id, "Основная группа", "1-2 дня");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Ставрополь", "1-2 дня");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Склады 2 дня", "3-4 дня");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Склады 7 дней", "8-9 дней");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            id = providers.AddProvider("4точки", 2);
            sw.Stop();
            Console.WriteLine("Добавили поставщика " + id + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Краснодар", "1-3 дня");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ids = providers.AddStock(id, "Москва", "7-9 дней");
            sw.Stop();
            Console.WriteLine("Добавили склад " + id + "-" + ids + " " + sw.ElapsedMilliseconds + " мс");

            ids = providers.AddStock(id, "Домодедово", "7-9 дней");
            ids = providers.AddStock(id, "Москва склад 2", "7-9 дней");
            ids = providers.AddStock(id, "Москва склад 3", "7-9 дней");
            ids = providers.AddStock(id, "Москва склад 4", "7-9 дней");

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
            EParsingParam parsingParam1 = new EParsingParam(fileName1, "A0");
            parsingParam1.AddStringVal("более 20", 20);
            parsingParam1.AddStringVal("более 40", 40);
            ESheet eSheet = new ESheet(null, "D");
            string[] bufIndex1 = { "B" };
            eSheet.AddBufIndex(bufIndex1);
            eSheet.AddCountIndex("A0", "E");
            eSheet.AddCountIndex("A1", "F");
            eSheet.AddCountIndex("A2", "G");
            eSheet.AddCountIndex("A3", "H");
            parsingParam1.Add(eSheet);

            eParsingParams.Add(parsingParam1);

            /*
            parsingParam1 = new EParsingParam(fileName2, id);
            parsingParam1.AddStringVal("более 20", 20);
            parsingParam1.AddStringVal("более 40", 40);
            string[] bufIndex2 = { "B", "D" };
            eSheet = new ESheet("1 Шины (Краснодар Индустриальны", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A0", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("3 Шины (Москва)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A1", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("4 Шины (Домодедово)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A2", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("5 Шины (Склад 2)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A3", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("6 Шины (Склад 3)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A4", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("7 Шины (Склад 4)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A5", "S");
            parsingParam1.Add(eSheet);
            
            eParsingParams.Add(parsingParam1);
            */
            

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
            dBase.SaveDataOnlyProviderExcell(@"C:\Users\ACER\Desktop\Прайсы\Data.xlsx", new string[] { "A0" }, new ExcelDefaultOutParametrics() {
                IsIdPosition = false,
                IsNameProduct = false
            },
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
            Console.WriteLine("Подготовка данных для 2-го парсинга (только для A0-A1)" + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            List<ParsingRow> parsings2 = EParsing.GetParsingRows(parsingParam2, out warning);
            sw.Stop();
            Console.WriteLine("Парсинг файла xlsx " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            ParsingRow[] rows2 = dictionary.Analysis(ref parsings2).Where(x => x.Resault is GResault).ToArray();
            sw.Stop();
            Console.WriteLine("Анализ данных " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.AddRows(rows2);
            sw.Stop();
            Console.WriteLine("Добавление в базу(оперативка) " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.Save();
            sw.Stop();
            Console.WriteLine("Сохранение базы XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.SaveDataExcel(@"C:\Users\ACER\Desktop\Прайсы\Data2.xlsx", new string[] { "A0" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data2.xlsx' " + sw.ElapsedMilliseconds + " мс");

            dictionary.Save();

            dictionary.Close();

            Console.WriteLine("готово");
            Console.WriteLine("********************************************************************\n\n\n");
            Console.ReadKey();
        }

        private static void Test2()
        {
            Console.WriteLine("********************************************************************\n");
            Console.WriteLine("-------------------------TEST2--------------------------------------\n");

            var fileName1 = @"C:\Users\ACER\Desktop\Прайсы\ВячеСлавик.xlsx";
            var fileName2 = @"C:\Users\ACER\Desktop\Прайсы\4точки.xlsx";

            Stopwatch sw = new Stopwatch();
            sw.Start();
            Providers providers = new Providers();

            sw.Stop();
            Console.WriteLine("Инициализация поставщиков XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            List<EParsingParam> eParsingParams = new List<EParsingParam>();
            EParsingParam parsingParam1 = new EParsingParam(fileName1, "A0");
            ESheet eSheet = new ESheet(null, "D");
            string[] bufIndex1 = { "B" };
            eSheet.AddBufIndex(bufIndex1);
            eSheet.AddCountIndex("A0", "E");
            eSheet.AddCountIndex("A1", "F");
            eSheet.AddCountIndex("A2", "G");
            eSheet.AddCountIndex("A3", "H");
            parsingParam1.Add(eSheet);

            eParsingParams.Add(parsingParam1);

            parsingParam1 = new EParsingParam(fileName2, "A1");
            string[] bufIndex2 = { "B", "D" };
            eSheet = new ESheet("3 Шины (Москва)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A1", "S");
            parsingParam1.Add(eSheet);
            eParsingParams.Add(parsingParam1);

            sw.Stop();
            Console.WriteLine("Подготовили параметры парсинга " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();

            List<ParsingRow> parsings = EParsing.GetParsingRows(eParsingParams.ToArray(), out List<string> warning);

            sw.Stop();
            Console.WriteLine("Парсинг файлов xlsx " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            Dictionary dictionary = new Dictionary();
            sw.Stop();
            Console.WriteLine("Инициализация словаря XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            var res = dictionary.Analysis(ref parsings);
            ParsingRow[] rows = res.Where(x => x.Resault is GResault).ToArray();
            ParsingRow[] bRows = res.Where(x => x.Resault is BResault).ToArray();
            ParsingRow[] NRows = res.Where(x => x.Resault is NResault).ToArray();
            sw.Stop();
            Console.WriteLine("Анализ данных " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            DBase dBase = new DBase();
            sw.Stop();
            Console.WriteLine("Инициализация базы XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.AddRows(rows);
            sw.Stop();
            Console.WriteLine("Добавление в базу(оперативка) " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.Save();
            sw.Stop();
            Console.WriteLine("Сохранение базы XML " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.SaveDataExcel(@"C:\Users\ACER\Desktop\Прайсы\Data4.xlsx", new string[] { "A1" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data4.xlsx' " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();

            dBase.SaveDataOnlyProviderExcell(@"C:\Users\ACER\Desktop\Прайсы\Data5.xlsx", new string[] { "A0" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data5.xlsx' " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            dBase.SaveDataExcel(@"C:\Users\ACER\Desktop\Прайсы\Data6.xlsx", new string[] { "A0", "A1" }, new ExcelDefaultOutParametrics(false, false, false
                , false, false, false, true, true,true), new ExcelProviderOutParametrics(providers, true, false, false, true), new ExcelProductOutParametrics(dictionary, true, false,
                false, false, false, true, false, false, true, false, false, true, true, true, true, true, false, false, false, false, false, true, false,true,true));
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data6.xlsx' " + sw.ElapsedMilliseconds + " мс");

            dictionary.Save();

            dictionary.Close();

            Console.WriteLine("готово");
            Console.WriteLine("********************************************************************\n\n\n");
            Console.ReadKey();
        }
    }
}