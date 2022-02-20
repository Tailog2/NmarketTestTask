using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using NmarketTestTask.Models;
using NmarketTestTask.Parsers;

namespace NmarketTestTask
{
    class Program
    {
        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof(List<House>));

        static void Main(string[] args)
        {
            Directory.CreateDirectory("Result");

            ParseAllFiles(new ExcelParser(), @"Files\Excel");
            ParseAllFiles(new HtmlParser(), @"Files\Html");
        }

        private static void ParseAllFiles(IParser parser, string folder)
        {
            foreach (var file in Directory.GetFiles(folder))
            {
                var result = parser.GetHouses(file);
                var resultFile = $@"Result\{Path.GetFileName(file)}.xml";

                using (var sw = new StreamWriter(resultFile))
                    Serializer.Serialize(sw, result);
            }
        }
    }
}
