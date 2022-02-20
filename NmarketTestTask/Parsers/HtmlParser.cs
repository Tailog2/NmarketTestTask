using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HtmlAgilityPack;
using NmarketTestTask.Models;

namespace NmarketTestTask.Parsers
{
    public class HtmlParser : IParser
    {
        private List<House> houses;

        public IList<House> GetHouses(string path)
        {
            houses = new List<House>();

            var doc = new HtmlDocument();
            doc.Load(path, Encoding.GetEncoding("windows-1251"), false);

            // Извлекает таблицу с данными без заголовков
            var tableData = doc.DocumentNode.SelectSingleNode("//tbody").SelectNodes(".//tr");
            ProcessRows(tableData);

            return houses;
        }

        // Проходит через каждую строку в таблице
        // Для каждой строчки извлекает имя дома используя имя класса СSS и каритру
        // Также для каждой строчки извлекает заполненный данными объект класса Flat
        // Добавляет квартиру в дом с именем извлеченным ранее
        private void ProcessRows(HtmlNodeCollection tableData)
        {
            foreach (var row in tableData)
            {
                var houseName = row.SelectSingleNode(".//td[@class='house']").InnerText;
                var flat = GetFlat(row);                
                AddFlat(flat, houseName);
            }
        }

        // Принимает строку из таблицы 
        // Извлекает данные о цене и номере квартиры используя имена класcов СSS
        // Возвращает заполненный данными объект Flat
        private Flat GetFlat(HtmlNode row)
        {
            Flat flat = new Flat()
            {
                Number = row.SelectSingleNode(".//td[@class='number']").InnerText,
                Price = row.SelectSingleNode(".//td[@class='price']").InnerText
            };

            return flat;
        }

        // Принимает заполненный данными объект Flat и имя дома
        // Если имя дома отсутсвиет в переменной houses создает новый объект класса House и добовляет его
        // Если имя дома присутсвует в переменной houses, находит дом и добавляет объект Flaт
        private void AddFlat(Flat flat, string houseName)
        {
            if (houses.Any(h => h.Name == houseName))
            {
                houses.Find(h => h.Name == houseName).Flats.Add(flat);
            }
            else
            {
                houses.Add(new House() { Name = houseName, Flats = new List<Flat>() });
                houses.Find(h => h.Name == houseName).Flats.Add(flat);
            }

            
        }
    }
}
