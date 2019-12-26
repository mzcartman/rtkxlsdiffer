using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace RtkXlsDiffer
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Инструкция:");
            Console.WriteLine();
            Console.WriteLine("1. Укажите путь к xls(x)-файлу со старыми данными.");
            Console.WriteLine("   Лист со списком должен быть ПЕРВЫМ и иметь следующие колонки:");
            Console.WriteLine("   Принадлежность, МРФ, Филиал, Код региона, Фирма-производитель оборудования, Модель оборудования, Имя устройства в DNS или его обозначение на схемах, Коммутатор, Наименование устройства/платы, Серийный номер, Версия ПО, Дата ввода в эксплуатацию, Дата начала гарантийного срока, Дата окончания гарантийного срока, Дата и № Контракта на поставку, Поставщик оборудования, Адрес установки оборудования, Детализация по полю Субъект, Функциональное назначение, Шасси/плата, IP-адрес устройства, Номер, Уровень сети"); 
            Console.WriteLine("   Порядок столбцов ВАЖЕН!");
            Console.WriteLine();
            Console.WriteLine("2. Укажите путь к xls(x)-файлу с новыми данными.");
            Console.WriteLine("   Все страницы с устройствами должны содержать следующие заголовки:");
            Console.WriteLine("   - ipaddress");
            Console.WriteLine("   - hostname");
            Console.WriteLine("   - model");
            Console.WriteLine("   - vendor");
            Console.WriteLine("   - commutator");
            Console.WriteLine("   - serial");
            Console.WriteLine("   - name");
            Console.WriteLine("   Порядок столбцов НЕ важен.");
            Console.WriteLine();
            Console.WriteLine("3. Подождите немного.");
            Console.WriteLine();
            Console.WriteLine("4. ????");
            Console.WriteLine();
            Console.WriteLine("5. PROFIT!!");
            Console.WriteLine();

            Console.WriteLine("Итак, введите путь к файлу со старыми устройствами и нажмите ENTER:");
            string oldpath = Console.ReadLine();
            Console.WriteLine();

            Console.WriteLine("Теперь введите путь к файлу с новыми устройствами и нажмите ENTER:");
            string newpath = Console.ReadLine();

#if DEBUG
            oldpath = "G:\\old.xlsx";
            newpath = "G:\\new.xlsx";
#endif

            Console.Clear();
            bool ipaddr_check = false;
            bool hostname_check = false;
            bool model_check = false;
            bool vendor_check = false;
            bool commutator_check = false;
            bool serial_check = false;
            bool naimenovanie_check = false;

            Application newdevxls = new Application();
            Workbook newdevxls_wb = newdevxls.Workbooks.Open(newpath);

            Dictionary<string, PGSO> newDev = new Dictionary<string, PGSO>();
            try
            {
                int fullnewdevcnt = 0;

                for (int j = 1; j <= newdevxls_wb.Sheets.Count; j++)
                {
                    _Worksheet xlWorksheet = newdevxls_wb.Sheets[j];
                    Range xlRange = xlWorksheet.UsedRange;

                    object[,] indexMatrix = (object[,])xlRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                    int rowCount = indexMatrix.GetLength(0);
                    int colCount = indexMatrix.GetLength(1);

                    fullnewdevcnt += rowCount - 1;

                    int ipaddr_col = 0;
                    int hostname_col = 0;
                    int model_col = 0;
                    int vendor_col = 0;
                    int commutator_col = 0;
                    int serial_col = 0;
                    int naimenovanie_col = 0;

                    for (int i = 1; i <= colCount; i++)
                    {
                        if (indexMatrix[1, i] != null && indexMatrix[1, i] != null)
                        {
                            if (indexMatrix[1, i].ToString().ToLower().Trim() == "ipaddress") ipaddr_col = i;
                            if (indexMatrix[1, i].ToString().ToLower().Trim() == "hostname") hostname_col = i;
                            if (indexMatrix[1, i].ToString().ToLower().Trim() == "model") model_col = i;
                            if (indexMatrix[1, i].ToString().ToLower().Trim() == "vendor") vendor_col = i;
                            if (indexMatrix[1, i].ToString().ToLower().Trim() == "commutator") commutator_col = i;
                            if (indexMatrix[1, i].ToString().ToLower().Trim() == "serial") serial_col = i;
                            if (indexMatrix[1, i].ToString().ToLower().Trim() == "name") naimenovanie_col = i;
                        }
                    }

                    Console.SetCursorPosition(0, 15);
                    if (ipaddr_col == 0) Console.WriteLine("!!!! ipaddress column not found!");
                    if (hostname_col == 0) Console.WriteLine("!!!! hostname column not found!");
                    if (model_col == 0) Console.WriteLine("!!!! model column not found!");
                    if (vendor_col == 0) Console.WriteLine("!!!! vendor column not found!");
                    if (commutator_col == 0) Console.WriteLine("!!!! commutator column not found!");
                    if (serial_col == 0) Console.WriteLine("!!!! serial column not found!");
                    if (naimenovanie_col == 0) Console.WriteLine("!!!! name column not found!");

                     model_check = model_col!= 0;
                     commutator_check = commutator_col != 0;
                     serial_check = serial_col != 0;
                     naimenovanie_check = naimenovanie_col != 0;

                    for (int i = 2; i <= rowCount; i++) //taking care of each Row  
                    {
                        Console.SetCursorPosition(0, 0);
                        Console.WriteLine("Чтение списка новых устройств... " + i);

                        try
                        {
                            PGSO pgso = new PGSO();
                            pgso.ipaddr = (ipaddr_col != 0 && indexMatrix[i, ipaddr_col] != null && indexMatrix[i, ipaddr_col] != null) ? indexMatrix[i, ipaddr_col].ToString().Trim() : "";
                            pgso.hostname = (hostname_col != 0 && indexMatrix[i, hostname_col] != null && indexMatrix[i, hostname_col] != null) ? indexMatrix[i, hostname_col].ToString().Trim() : "";
                            pgso.model = (model_col != 0 && indexMatrix[i, model_col] != null && indexMatrix[i, model_col] != null) ? indexMatrix[i, model_col].ToString().Trim() : "";
                            pgso.vendor = (vendor_col != 0 && indexMatrix[i, vendor_col] != null && indexMatrix[i, vendor_col] != null) ? indexMatrix[i, vendor_col].ToString().Trim() : "";
                            pgso.commutator = (commutator_col != 0 && indexMatrix[i, commutator_col] != null && indexMatrix[i, commutator_col] != null) ? indexMatrix[i, commutator_col].ToString().Trim() : "";
                            pgso.serial = (serial_col != 0 && indexMatrix[i, serial_col] != null && indexMatrix[i, serial_col] != null) ? indexMatrix[i, serial_col].ToString().Trim() : "";
                            pgso.naimenovanie = (naimenovanie_col != 0 && indexMatrix[i, naimenovanie_col] != null && indexMatrix[i, naimenovanie_col] != null) ? indexMatrix[i, naimenovanie_col].ToString().Trim() : "";

                            string un = "";
                            un += serial_check ? pgso.serial.ToLower() : "";
                            un += naimenovanie_check ? pgso.naimenovanie.ToLower() : "";
                            un += commutator_check ? pgso.commutator.ToLower() : "";
                            un += model_check ? pgso.model.ToLower() : "";
                            //un = un.Replace(" ", "_").Replace("-", "_").Replace(".", "_").Replace(",", "_");

                            if (!newDev.ContainsKey(un)) newDev.Add(un, pgso);
                        }
                        catch { }
                    }

                    Console.WriteLine("Новых устройств в таблице: " + fullnewdevcnt);
                    Console.WriteLine("Загружено уникальных устройств: " + newDev.Count);

                    // xlWorkbook.Save();

                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //rule of thumb for releasing com objects:
                    //  never use two dots, all COM objects must be referenced and released individually
                    //  ex: [somthing].[something].[something] is bad

                    //release com objects to fully kill excel process from running in the background
                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);
                }

                //close and release
                newdevxls_wb.Close();
                Marshal.ReleaseComObject(newdevxls_wb);

                //quit and release
                newdevxls.Quit();
                Marshal.ReleaseComObject(newdevxls);
                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.Clear();
                Console.WriteLine("Well, something went wrong. There is a error:" + Environment.NewLine + ex.ToString());
                newdevxls_wb.Close(0);
                newdevxls.Quit();

                Marshal.ReleaseComObject(newdevxls_wb);
                Marshal.ReleaseComObject(newdevxls);
                Marshal.FinalReleaseComObject(newdevxls);
                goto errr;
            }

            Application olddevxls = new Application();
            Workbook olddevxls_wb = olddevxls.Workbooks.Open(oldpath);

            Dictionary<string, PGSO> oldDev = new Dictionary<string, PGSO>();

            try
            {
                _Worksheet xlWorksheet = olddevxls_wb.Sheets[1];
                Range xlRange = xlWorksheet.UsedRange;

                object[,] indexMatrix = (object[,])xlRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                int rowCount = indexMatrix.GetLength(0);
                int colCount = indexMatrix.GetLength(1);

                for (int i = 1; i <= rowCount; i++) //taking care of each Row  
                {
                    Console.SetCursorPosition(0, 3);
                    Console.WriteLine("Читается список старых устройств... " + i);

                    PGSO pgso = new PGSO();

                    pgso.prinadlejn = (indexMatrix[i, 1] != null && indexMatrix[i, 1] != null) ? indexMatrix[i, 1].ToString().Trim() : "";
                    pgso.mrf = (indexMatrix[i, 2] != null && indexMatrix[i, 2] != null) ? indexMatrix[i, 2].ToString().Trim() : "";
                    pgso.filial = (indexMatrix[i, 3] != null && indexMatrix[i, 3] != null) ? indexMatrix[i, 3].ToString().Trim() : "";
                    pgso.codreg = (indexMatrix[i, 4] != null && indexMatrix[i, 4] != null) ? indexMatrix[i, 4].ToString().Trim() : "";
                    pgso.vendor = (indexMatrix[i, 5] != null && indexMatrix[i, 5] != null) ? indexMatrix[i, 5].ToString().Trim() : "";
                    pgso.model = (indexMatrix[i, 6] != null && indexMatrix[i, 6] != null) ? indexMatrix[i, 6].ToString().Trim() : "";
                    pgso.hostname = (indexMatrix[i, 7] != null && indexMatrix[i, 7] != null) ? indexMatrix[i, 7].ToString().Trim() : "";
                    pgso.commutator = (indexMatrix[i, 8] != null && indexMatrix[i, 8] != null) ? indexMatrix[i, 8].ToString().Trim() : "";
                    pgso.naimenovanie = (indexMatrix[i, 9] != null && indexMatrix[i, 9] != null) ? indexMatrix[i, 9].ToString().Trim() : "";
                    pgso.serial = (indexMatrix[i, 10] != null && indexMatrix[i, 10] != null) ? indexMatrix[i, 10].ToString().Trim() : "";
                    pgso.pover = (indexMatrix[i, 11] != null && indexMatrix[i, 11] != null) ? indexMatrix[i, 11].ToString().Trim() : "";                    
                    pgso.startdate = (indexMatrix[i, 12] != null && indexMatrix[i, 12] != null) ? indexMatrix[i, 12].ToString().Trim() : "";
                    pgso.garantstart = (indexMatrix[i, 13] != null && indexMatrix[i, 13] != null) ? indexMatrix[i, 13].ToString().Trim() : "";
                    pgso.garantend = (indexMatrix[i, 14] != null && indexMatrix[i, 14] != null) ? indexMatrix[i, 14].ToString().Trim() : "";
                    pgso.contractinfo = (indexMatrix[i, 15] != null && indexMatrix[i, 15] != null) ? indexMatrix[i, 15].ToString().Trim() : "";
                    pgso.postavshik = (indexMatrix[i, 16] != null && indexMatrix[i, 16] != null) ? indexMatrix[i, 16].ToString().Trim() : "";
                    pgso.adres = (indexMatrix[i, 17] != null && indexMatrix[i, 17] != null) ? indexMatrix[i, 17].ToString().Trim() : "";
                    pgso.detail = (indexMatrix[i, 18] != null && indexMatrix[i, 18] != null) ? indexMatrix[i, 18].ToString().Trim() : "";
                    pgso.function = (indexMatrix[i, 19] != null && indexMatrix[i, 19] != null) ? indexMatrix[i, 19].ToString().Trim() : "";
                    pgso.chasis = (indexMatrix[i, 20] != null && indexMatrix[i, 20] != null) ? indexMatrix[i, 20].ToString().Trim() : "";
                    pgso.ipaddr = (indexMatrix[i, 21] != null && indexMatrix[i, 21] != null) ? indexMatrix[i, 21].ToString().Trim() : "";
                    pgso.asnum = (indexMatrix[i, 22] != null && indexMatrix[i, 22] != null) ? indexMatrix[i, 22].ToString().Trim() : "";
                    pgso.netlevel = (indexMatrix[i, 23] != null && indexMatrix[i, 23] != null) ? indexMatrix[i, 23].ToString().Trim() : "";

                    string un = "";
                    un += serial_check ? pgso.serial.ToLower() : "";
                    un += naimenovanie_check ? pgso.naimenovanie.ToLower() : "";
                    un += commutator_check ? pgso.commutator.ToLower() : "";
                    un += model_check ? pgso.model.ToLower() : "";
                    //un = un.Replace(" ", "_").Replace("-", "_").Replace(".", "_").Replace(",", "_");

                    if (!oldDev.ContainsKey(un)) oldDev.Add(un, pgso);
                }

                Console.WriteLine("Старых устройств в таблице: " + (rowCount - 1));
                Console.WriteLine("Загружено уникальных устройств: " + (oldDev.Count - 1));

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                olddevxls_wb.Close();
                Marshal.ReleaseComObject(olddevxls_wb);

                //quit and release
                olddevxls.Quit();
                Marshal.ReleaseComObject(olddevxls);
            }
            catch (Exception ex)
            {
                Console.Clear();
                Console.WriteLine("Well, something went wrong. There is a error:" + Environment.NewLine + ex.ToString());
                olddevxls_wb.Close(0);
                olddevxls.Quit();

                Marshal.ReleaseComObject(olddevxls_wb);
                Marshal.ReleaseComObject(olddevxls);
                Marshal.FinalReleaseComObject(olddevxls);

                goto errr;
            }            

            int newDevs = 0;

            string check = "Сравнение будет основано на полях ";
            check += serial_check ? "serial, " : "";
            check += naimenovanie_check ?"name, " : "";
            check += commutator_check ? "commutator, " : "";
            check += model_check ? "model, " : "";
            check = check.Trim().TrimEnd(',');

            Console.WriteLine();
            Console.WriteLine("---------------" + check + "---------------");
            Console.WriteLine();

            foreach (var dev in newDev)
            {
                Console.SetCursorPosition(0, 9);
                Console.WriteLine("Добавляем новые устройства... " + newDevs + "   ");

                if (!oldDev.ContainsKey(dev.Key))
                {
                    oldDev.Add(dev.Key, dev.Value);
                    newDevs++;
                }
            }

            Console.SetCursorPosition(0, 9);
            Console.WriteLine("Добавлено новых устройств: " + newDevs + "    ");
            Console.WriteLine("Общее количество устройств в новом списке: " + (oldDev.Count - 1) + "");
            Console.WriteLine("Нажмите Enter для создания новой таблицы...");
            Console.ReadLine();
            Console.WriteLine("Создание таблицы...");

            Application newdoc = new Application();
            Workbook newdoc_workBook = newdoc.Workbooks.Add(Type.Missing);
            Worksheet newdoc_workSheet = (Worksheet)newdoc_workBook.ActiveSheet;

            try
            {
                newdoc_workSheet.Name = "pgso";
                object[,] indexMatrix = new object[oldDev.Count, 24];

                int index = 0;

                foreach (var dev in oldDev)
                {
                    var pgsodev = dev.Value.ToArray();
                    for (int i = 0; i < 24; i++) indexMatrix[index, i] = pgsodev[i];
                    index++;
                }

                int rowCount = indexMatrix.GetLength(0);
                int columnCount = indexMatrix.GetLength(1);
                // Get an Excel Range of the same dimensions
                Range range = (Range)newdoc_workSheet.Cells[1, 1];
                range = range.get_Resize(rowCount, columnCount);
                // Assign the 2-d array to the Excel Range
                range.set_Value(XlRangeValueDataType.xlRangeValueDefault, indexMatrix);
                range.EntireColumn.NumberFormat = "@";
                range.EntireColumn.AutoFit();
                newdoc.Visible = true;
                newdoc.UserControl = true;
            }
            catch (Exception ex)
            {
                Console.Clear();
                Console.WriteLine("Well, something went wrong. There is a error:" + Environment.NewLine + ex.ToString());
                newdoc_workBook.Close(0);
                newdoc.Quit();

                Marshal.ReleaseComObject(newdoc_workBook);
                Marshal.ReleaseComObject(newdoc);
                Marshal.FinalReleaseComObject(newdoc);

                goto errr;
            }

            goto noerrr;

        errr:
            Console.ReadLine();
        noerrr:
            Console.WriteLine("Done!");
        }
    }

    public class PGSO
    {
        public string prinadlejn;
        public string mrf;
        public string filial;
        public string codreg;
        public string vendor;
        public string model;
        public string hostname;
        public string commutator;
        public string naimenovanie;
        public string serial;
        public string pover;
        public string startdate;
        public string garantstart;
        public string garantend;
        public string contractinfo;
        public string postavshik;
        public string adres;
        public string detail;
        public string function;
        public string chasis;
        public string ipaddr;
        public string asnum;
        public string netlevel;

        public string[] ToArray()
        {
            string[] array = new string[25];

            array[00] = prinadlejn;
            array[01] = mrf;
            array[02] = filial;
            array[03] = codreg;
            array[04] = vendor;
            array[05] = model;
            array[06] = hostname;
            array[07] = commutator;
            array[08] = naimenovanie;
            array[09] = serial;
            array[10] = pover;
            array[12] = startdate;
            array[13] = garantstart;
            array[14] = garantend;
            array[15] = contractinfo;
            array[16] = postavshik;
            array[17] = adres;
            array[18] = detail;
            array[19] = function;
            array[20] = chasis;
            array[21] = ipaddr;
            array[22] = asnum;
            array[23] = netlevel;

            return array;
        }

        /*
            Принадлежность
            МРФ
            Филиал 
            Код региона
            Фирма-производитель оборудования
            Модель оборудования
            Имя устройства в DNS или его обозначение на схемах
            Коммутатор
            Наименование устройства/платы
            Серийный номер
            Версия ПО            
            Дата ввода в эксплуатацию
            Дата начала гарантийного срока
            Дата окончания гарантийного срока
            Дата и № Контракта на поставку
            Поставщик оборудования
            Адрес установки оборудования
            Детализация по полю Субъект	
            Функциональное назначение	
            Шасси/плата	
            IP-адрес устройства	
            Номер	
            Уровень сети
        */

    }
}