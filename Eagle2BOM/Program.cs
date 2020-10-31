using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace Eagle2BOM_PCL
{
    class Program
    {
        struct PartStruct
        {
            public string Part;
            public string Value;
            public string Package;
            public string Library;
            public PointF Position;
            public string Orientation;
        };

        static float convertToMillimeters = 1.0f;

        static void Setup()
        {
            // Force decimal point
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB");
        }

        static List<string> ConvertFixedLengthString2StringArray(string fixedLenString)
        {
            var lengths = new[] { 10, 17, 33, 21, 23, 9999 };

            var totalLength = fixedLenString.Length;
            var stringArray = new List<string>();
            var i = 0;
            foreach (var length in lengths)
            {
                stringArray.Add(fixedLenString.Substring(i, (length > totalLength) ? totalLength - i : length - 1).TrimEnd());
                i += length - 1;
            }

            return stringArray;
        }

        static List<PartStruct> ConvertStringArrayToPartList(string[] stringArray)
        {
            var partList = new List<PartStruct>();

            foreach (var partLine in stringArray)
            {
                var parts = ConvertFixedLengthString2StringArray(partLine);

                // parse the position from (x, y) to x,y
                var position = parts[4].Trim(new char[] { '(', ')' });
                var xy = position.Split(' ');

                var part = new PartStruct();
                part.Part = parts[0];
                part.Value = parts[1];
                part.Package = parts[2];
                part.Library = parts[3];
                part.Position.X = float.Parse(xy[0]) * convertToMillimeters;
                part.Position.Y = float.Parse(xy[1]) * convertToMillimeters;
                part.Orientation = parts[5].Replace('R', '-');
                part.Orientation = part.Orientation.Replace('L', '+');

                partList.Add(part);
            }

            return partList;
        }

        static List<PartStruct> ReadPartlist(string filename)
        {
            if (!File.Exists(filename))
                return null;

            var partLines = File.ReadAllLines(filename);

            // Remove blank lines
            partLines = partLines.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            partLines = partLines.Skip(4).ToArray();

            var labels = ConvertFixedLengthString2StringArray(partLines.FirstOrDefault());

            if (labels.Where(x => x.Contains("inch")).ToList().Count > 0)
                convertToMillimeters = 25.4f;
            partLines = partLines.Skip(1).ToArray();

            return ConvertStringArrayToPartList(partLines);
        }

        static void Main(string[] args)
        {
            if (args.Length != 2)
                return; // expect 2 arguments

            // both files must exist
            foreach (var filename in args)
                if (!File.Exists(filename))
                    return;

            Setup();

            var PartList = ReadPartlist(args[0]);

            //----------------------------------------------------------------------------------------
            // Lookup
            var PartLookup = new Dictionary<string, string>();

            var partsList = File.ReadAllLines(args[1]);
            // Remove blank lines
            partsList = partsList.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            partsList = partsList.Skip(1).ToArray();

            foreach (var line in partsList)
            {
                var parts = line.Split(',');
                PartLookup.Add(parts[0] + "," + parts[1], parts[2]);
            }

            //----------------------------------------------------------------------------------------
            // Construct BOM
            var bomFilename = args[0] + "_bom.csv";

            var sw = new StreamWriter(bomFilename);
            sw.WriteLine("Comment,Designator,Footprint,PartNr");

            var bomList = PartList.GroupBy(p => new { p.Value })
                                  .Select(g => g.First())
                                  .ToList();

            foreach (var t in bomList)
            {
                if (string.IsNullOrEmpty(t.Value))
                    continue;

                PartLookup.TryGetValue(t.Value + "," + t.Package, out var manuPart);

                sw.Write($"{t.Value},\"");

                var p = PartList.Where(y => (y.Value == t.Value)).ToList();
                foreach (var item in p)
                    sw.Write($"{item.Part},");

                sw.WriteLine($"\",{t.Package},{manuPart}");
            }
            sw.Close();

            //----------------------------------------------------------------------------------------
            // Construct CPL

            var cplFilename = args[0] + "_cpl.csv";

            sw = new StreamWriter(cplFilename);
            sw.WriteLine("Designator,Mid X,Mid Y,Layer,Rotation");

            foreach (var t in PartList)
                sw.WriteLine($"{t.Part},{t.Position.X}mm,{t.Position.Y}mm,{t.Orientation}");

            sw.Close();

        }
    }
}
