using ProjectX.DataBase;
using ProjectX.Dict;
using ProjectX.ExcelParsing;
using ProjectX.Information;
using ProjectX.Loger;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ProjectX
{
    class Program
    {


        public static void Main(string[] args)
        {


            Test1();

            Test2();

        }

        static void Test1() {

            Console.WriteLine("********************************************************************\n");
            Console.WriteLine("-------------------------TEST1--------------------------------------\n");

            var fileName1 = @"C:\Users\ACER\Desktop\Прайсы\ВячеСлавик.xlsx";
            var fileName2 = @"C:\Users\ACER\Desktop\Прайсы\4точки.xlsx";

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
            EParsingParam parsingParam1 = new EParsingParam(fileName1, "A1");
            ESheet eSheet = new ESheet(null, "D");
            string[] bufIndex1 = { "B" };
            eSheet.AddBufIndex(bufIndex1);
            eSheet.AddCountIndex("A1", "E");
            eSheet.AddCountIndex("A2", "F");
            eSheet.AddCountIndex("A3", "G");
            eSheet.AddCountIndex("A4", "H");
            parsingParam1.Add(eSheet);

            eParsingParams.Add(parsingParam1);


            parsingParam1 = new EParsingParam(fileName2, id);
            string[] bufIndex2 = { "B", "D" };
            eSheet = new ESheet("1 Шины (Краснодар Индустриальны", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A1", "S");
            parsingParam1.Add(eSheet);
            eSheet = new ESheet("3 Шины (Москва)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A2", "S");
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
            ParsingRow[] rows = dictionary.Analysis(ref parsings).Where(x => x.Resault is GResault).ToArray();
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
            dBase.SaveDataExcel(@"C:\Users\ACER\Desktop\Прайсы\Data.xlsx", new string[] { "A1", "A2" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data.xlsx' " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();

            dBase.SaveDataOnlyProviderExcell(@"C:\Users\ACER\Desktop\Прайсы\Data3.xlsx", new string[] { "A1", "A2" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data3.xlsx' " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();
            EParsingParam parsingParam2 = new EParsingParam(fileName1, "A1");
            ESheet eSheet2 = new ESheet(null, "D");
            eSheet2.AddBufIndex(bufIndex1);
            eSheet2.AddCountIndex("A2", "F");
            parsingParam2.Add(eSheet2);
            sw.Stop();
            Console.WriteLine("Подготовка данных для 2-го парсинга (только для A1-A2)" + sw.ElapsedMilliseconds + " мс");

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
            dBase.SaveDataExcel(@"C:\Users\ACER\Desktop\Прайсы\Data2.xlsx", new string[] { "A1" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data2.xlsx' " + sw.ElapsedMilliseconds + " мс");


            dictionary.Save();

            dictionary.Close();

            Console.WriteLine("готово");
            Console.WriteLine("********************************************************************\n\n\n");
            Console.ReadKey();

        }

        static void Test2() {

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
            EParsingParam parsingParam1 = new EParsingParam(fileName1, "A1");
            ESheet eSheet = new ESheet(null, "D");
            string[] bufIndex1 = { "B" };
            eSheet.AddBufIndex(bufIndex1);
            eSheet.AddCountIndex("A1", "E");
            eSheet.AddCountIndex("A2", "F");
            eSheet.AddCountIndex("A3", "G");
            eSheet.AddCountIndex("A4", "H");
            parsingParam1.Add(eSheet);

            eParsingParams.Add(parsingParam1);


            parsingParam1 = new EParsingParam(fileName2, "A2");
            string[] bufIndex2 = { "B", "D" };
            eSheet = new ESheet("3 Шины (Москва)", "U");
            eSheet.AddBufIndex(bufIndex2);
            eSheet.AddCountIndex("A2", "S");
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
            ParsingRow[] rows = dictionary.Analysis(ref parsings).Where(x => x.Resault is GResault).ToArray();
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
            dBase.SaveDataExcel(@"C:\Users\ACER\Desktop\Прайсы\Data4.xlsx", new string[] { "A2" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data4.xlsx' " + sw.ElapsedMilliseconds + " мс");

            sw.Restart();

            dBase.SaveDataOnlyProviderExcell(@"C:\Users\ACER\Desktop\Прайсы\Data5.xlsx", new string[] { "A1" });
            sw.Stop();
            Console.WriteLine("Сохранение базы в xlsx 'Data5.xlsx' " + sw.ElapsedMilliseconds + " мс");


            sw.Restart();
            dBase.SaveDataExcel(@"C:\Users\ACER\Desktop\Прайсы\Data6.xlsx", new string[] { "A1" },new ExcelDefaultOutParametrics(false,false,false
                ,false,false,false,true,true),new ExcelProviderOutParametrics(providers,true,false,false,true),new ExcelProductOutParametrics(dictionary,true,false,
                false,false,false,true,false,false,false,false,false,true,true,true,false,false,false,false,false,false,false,true,false));
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
