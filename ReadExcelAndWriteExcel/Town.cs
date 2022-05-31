using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelAndWriteExcel
{
    public class Town
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string County { get; set; }
        public string District { get; set; }
        public int Area { get; set; }
        public int Population { get; set; }
        public int ApartmentsCount { get; set; }
        public Town(string line)
        {
            string[] lineData = line.Split(';');
            Name = lineData[0];
            Type = lineData[1];
            County = lineData[2];
            District = lineData[3];
            Area = int.Parse(lineData[4]);
            Population = int.Parse(lineData[5]);
            ApartmentsCount = int.Parse(lineData[6].Substring(0, lineData[6].Length - 1));
        }
        public override string ToString()
        {
            return string.Format($"{Name};{Type};{County};{District};{Area};{Population};{ApartmentsCount}");
        }
    }
}
