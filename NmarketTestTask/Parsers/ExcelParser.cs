using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using NmarketTestTask.Models;

namespace NmarketTestTask.Parsers
{
    public class ExcelParser : IParser
    {
        private IXLWorksheet sheet;
        private List<House> houses;

        public IList<House> GetHouses(string path)
        {
            houses = new List<House>();

            try
            {
                var workbook = new XLWorkbook(path);
                sheet = workbook.Worksheets.First();
            }
            catch (Exception)
            {
                throw new Exception("Error during the file loading");
            }                

            // Извлекает ячейки с именами домов
            var housesInSheet = sheet.Cells().Where(c => c.GetValue<string>().Contains("Дом")).ToList();

            if (housesInSheet.Count == 0)
                throw new Exception("The program was not able to find any houses in the table");

            AddHouses(housesInSheet);

            return houses;
        }

        // Принимает список с ячейками домов
        // Создает новый обект класса House и заполняет его данными
        // Добавляет новый обект класса House в список houses
        private void AddHouses(List<IXLCell> housesInSheet)
        {
            for (int i = 0; i < housesInSheet.Count; i++)
            {
                var house = GetHouse(housesInSheet, i);
                houses.Add(house);
            }
        }

        // Принимает список с ячейками домов и его индекс в списке    
        // Извлекает ячейку с домом и  ячейку со следующем в списке домом
        // Если это последний дом в списке присваивает ячейке со следующим домом значение null
        // Извлекает все картины для дома соответствующего принятому индекс
        // Создает объект House, заполняет его данными и возращает его
        private House GetHouse(List<IXLCell> housesInSheet, int houseIndex)
        {

            var houseInSheet = housesInSheet[houseIndex];
            IXLCell nextHouseInSheet = null;

            if (houseIndex < housesInSheet.Count - 1)
                nextHouseInSheet = housesInSheet[houseIndex + 1];

            var flats = GetFlats(houseInSheet, nextHouseInSheet);

            var house = new House()
            {
                Name = houseInSheet.GetValue<string>(),
                Flats = flats
            };

            return house;
        }

        // Принимает ячейку с текущим домом и со следующим домом
        // Извлекает и возвращает все квартиры для текущего дома
        private List<Flat> GetFlats(IXLCell houseInSheet, IXLCell nextHouseInSheet)
        {
            List<IXLCell> flatsInSheet = FindFlats(houseInSheet, nextHouseInSheet);
            
            return ConvertFlatsCellsToFlats(flatsInSheet);
        }

        // Принимает ячейку с текущим домом и со следующим домом
        // Находит все квартиры для текущим домом
        // Находятся все квартиры начиная от ячейки текущим домом заканчивая ячейкой следующего дома
        // Поиск идет сверху вниз построчно
        // Возвращает список с ячейками квартир для текущего дома
        private List<IXLCell> FindFlats(IXLCell houseInSheet, IXLCell nextHouseInSheet)
        {
            List<IXLCell> flatsInSheet;

            if (nextHouseInSheet != null)
            {
                flatsInSheet = sheet.Cells()
                .Where(c => c.GetValue<string>().Contains("№"))
                .SkipWhile((cell, index) => cell.WorksheetRow().RowNumber() < houseInSheet.WorksheetRow().RowNumber())
                .TakeWhile((cell, index) => cell.WorksheetRow().RowNumber() < nextHouseInSheet.WorksheetRow().RowNumber())
                .ToList();
            }
            else
            {
                flatsInSheet = sheet.Cells()
                .SkipWhile((cell, index) => cell.WorksheetRow().RowNumber() < houseInSheet.WorksheetRow().RowNumber())
                .Where(c => c.GetValue<string>().Contains("№"))
                .ToList();
            }

            return flatsInSheet;
        }

        // Преобразует список с ячейками квартир в список с объектами класса Flat
        private List<Flat> ConvertFlatsCellsToFlats(List<IXLCell> flatsInSheet)
        {
            var flats = new List<Flat>();
            foreach (var flatInSheet in flatsInSheet)
            {
                var flatNumber = GetFlatNumber(flatInSheet);
                var price = GetPrice(flatInSheet);
                flats.Add(new Flat() { Number = flatNumber, Price = price });
            }
            return flats;
        }

        // Извлекает и возвращает номер квартиры
        private string GetFlatNumber(IXLCell flatInSheet)
        {
            var flatNumber = Regex.Match(flatInSheet.GetValue<string>(), @"\d+").Value;
            return flatNumber;
        }

        // Извлекает и возвращает цену квартиры
        private string GetPrice(IXLCell flatInSheet)
        {
            var price = sheet.Cell(flatInSheet.WorksheetRow().RowNumber() + 1, flatInSheet.WorksheetColumn().ColumnNumber()).GetValue<string>();
            return price;
        }
    }
}
