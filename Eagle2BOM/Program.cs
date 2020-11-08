using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Eagle2BOM_PCL
{
    public static class MyExtensions
    {
        public static string Mid(this string s, int start, int length)
        {
            return (s.Length > start + length) ?
                 s.Substring(start, length) :
                 s.Substring(start, s.Length - start);
        }
    }

    class Program
    {
        #region Signitures imported from http://pinvoke.net

        [DllImport("shfolder.dll", CharSet = CharSet.Auto)]
        internal static extern int SHGetFolderPath(IntPtr hwndOwner, int nFolder, IntPtr hToken, int dwFlags, StringBuilder lpszPath);

        [Flags()]
        enum SLGP_FLAGS
        {
            /// <summary>Retrieves the standard short (8.3 format) file name</summary>
            SLGP_SHORTPATH = 0x1,
            /// <summary>Retrieves the Universal Naming Convention (UNC) path name of the file</summary>
            SLGP_UNCPRIORITY = 0x2,
            /// <summary>Retrieves the raw path name. A raw path is something that might not exist and may include environment variables that need to be expanded</summary>
            SLGP_RAWPATH = 0x4
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        struct WIN32_FIND_DATAW
        {
            public uint dwFileAttributes;
            public long ftCreationTime;
            public long ftLastAccessTime;
            public long ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint dwReserved0;
            public uint dwReserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
            public string cAlternateFileName;
        }

        [Flags()]
        enum SLR_FLAGS
        {
            /// <summary>
            /// Do not display a dialog box if the link cannot be resolved. When SLR_NO_UI is set,
            /// the high-order word of fFlags can be set to a time-out value that specifies the
            /// maximum amount of time to be spent resolving the link. The function returns if the
            /// link cannot be resolved within the time-out duration. If the high-order word is set
            /// to zero, the time-out duration will be set to the default value of 3,000 milliseconds
            /// (3 seconds). To specify a value, set the high word of fFlags to the desired time-out
            /// duration, in milliseconds.
            /// </summary>
            SLR_NO_UI = 0x1,
            /// <summary>Obsolete and no longer used</summary>
            SLR_ANY_MATCH = 0x2,
            /// <summary>If the link object has changed, update its path and list of identifiers.
            /// If SLR_UPDATE is set, you do not need to call IPersistFile::IsDirty to determine
            /// whether or not the link object has changed.</summary>
            SLR_UPDATE = 0x4,
            /// <summary>Do not update the link information</summary>
            SLR_NOUPDATE = 0x8,
            /// <summary>Do not execute the search heuristics</summary>
            SLR_NOSEARCH = 0x10,
            /// <summary>Do not use distributed link tracking</summary>
            SLR_NOTRACK = 0x20,
            /// <summary>Disable distributed link tracking. By default, distributed link tracking tracks
            /// removable media across multiple devices based on the volume name. It also uses the
            /// Universal Naming Convention (UNC) path to track remote file systems whose drive letter
            /// has changed. Setting SLR_NOLINKINFO disables both types of tracking.</summary>
            SLR_NOLINKINFO = 0x40,
            /// <summary>Call the Microsoft Windows Installer</summary>
            SLR_INVOKE_MSI = 0x80
        }


        /// <summary>The IShellLink interface allows Shell links to be created, modified, and resolved</summary>
        [ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")]
        interface IShellLinkW
        {
            /// <summary>Retrieves the path and file name of a Shell link object</summary>
            void GetPath([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out WIN32_FIND_DATAW pfd, SLGP_FLAGS fFlags);
            /// <summary>Retrieves the list of item identifiers for a Shell link object</summary>
            void GetIDList(out IntPtr ppidl);
            /// <summary>Sets the pointer to an item identifier list (PIDL) for a Shell link object.</summary>
            void SetIDList(IntPtr pidl);
            /// <summary>Retrieves the description string for a Shell link object</summary>
            void GetDescription([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
            /// <summary>Sets the description for a Shell link object. The description can be any application-defined string</summary>
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            /// <summary>Retrieves the name of the working directory for a Shell link object</summary>
            void GetWorkingDirectory([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
            /// <summary>Sets the name of the working directory for a Shell link object</summary>
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            /// <summary>Retrieves the command-line arguments associated with a Shell link object</summary>
            void GetArguments([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
            /// <summary>Sets the command-line arguments for a Shell link object</summary>
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            /// <summary>Retrieves the hot key for a Shell link object</summary>
            void GetHotkey(out short pwHotkey);
            /// <summary>Sets a hot key for a Shell link object</summary>
            void SetHotkey(short wHotkey);
            /// <summary>Retrieves the show command for a Shell link object</summary>
            void GetShowCmd(out int piShowCmd);
            /// <summary>Sets the show command for a Shell link object. The show command sets the initial show state of the window.</summary>
            void SetShowCmd(int iShowCmd);
            /// <summary>Retrieves the location (path and index) of the icon for a Shell link object</summary>
            void GetIconLocation([Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath,
                int cchIconPath, out int piIcon);
            /// <summary>Sets the location (path and index) of the icon for a Shell link object</summary>
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
            /// <summary>Sets the relative path to the Shell link object</summary>
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
            /// <summary>Attempts to find the target of a Shell link, even if it has been moved or renamed</summary>
            void Resolve(IntPtr hwnd, SLR_FLAGS fFlags);
            /// <summary>Sets the path and file name of a Shell link object</summary>
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);

        }

        [ComImport, Guid("0000010c-0000-0000-c000-000000000046"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPersist
        {
            [PreserveSig]
            void GetClassID(out Guid pClassID);
        }


        [ComImport, Guid("0000010b-0000-0000-C000-000000000046"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPersistFile : IPersist
        {
            new void GetClassID(out Guid pClassID);
            [PreserveSig]
            int IsDirty();

            [PreserveSig]
            void Load([In, MarshalAs(UnmanagedType.LPWStr)]
            string pszFileName, uint dwMode);

            [PreserveSig]
            void Save([In, MarshalAs(UnmanagedType.LPWStr)] string pszFileName,
                [In, MarshalAs(UnmanagedType.Bool)] bool fRemember);

            [PreserveSig]
            void SaveCompleted([In, MarshalAs(UnmanagedType.LPWStr)] string pszFileName);

            [PreserveSig]
            void GetCurFile([In, MarshalAs(UnmanagedType.LPWStr)] string ppszFileName);
        }

        const uint STGM_READ = 0;
        const int MAX_PATH = 260;

        // CLSID_ShellLink from ShlGuid.h 
        [
            ComImport(),
            Guid("00021401-0000-0000-C000-000000000046")
        ]
        public class ShellLink
        {
        }

        #endregion


        public static string ResolveShortcut(string filename)
        {
            var link = new ShellLink();
            ((IPersistFile)link).Load(filename, STGM_READ);

            var sb = new StringBuilder(MAX_PATH);
            ((IShellLinkW)link).GetPath(sb, sb.Capacity, out var data, 0);

            var result = sb.ToString();

            return (string.IsNullOrEmpty(result)) ? filename : result;
        }

        struct PartStruct
        {
            public string Part;
            public string Value;
            public string Package;
            public string Library;
            public PointF Position;
            public int    Orientation;
        };

        static float convertToMillimeters = 1.0f;

        static Dictionary<string, int> ReadLegende(string legende)
        {
            if (legende.Contains("(inch)"))
                convertToMillimeters = 25.4f;

            legende = legende.Replace("(mm)",   "    ");
            legende = legende.Replace("(inch)", "      ");

            var dictionary = new Dictionary<string, int>();
            var start = 0;

            for (int i = 1; i < legende.Length; i++)
            {
                if (legende[i] != ' ' && ' ' == legende[i - 1])
                {
                    var key = legende.Substring(start, i - start).Trim();
                    if (!(key.Contains("(inch)") || (key.Contains("(mm)"))))
                        dictionary.Add(key, i - start);

                    start = i;
                }
            }
            dictionary.Add(legende.Substring(start, legende.Length - start).Trim(), legende.Length - start);

            return dictionary;
        }

        static Dictionary<string, string> ConvertFixedLengthString2StringArray(Dictionary<string, int> dictionary, string fixedLenString)
        {
            var elements = new Dictionary<string, string>();

            var start = 0;
            foreach (var item in dictionary)
            {
                var value = fixedLenString.Mid(start, item.Value).Trim();
                elements.Add(item.Key, value);

                start += item.Value;
            }

            return elements;
        }
        
        static List<PartStruct> ConvertStringArrayToPartList(Dictionary<string, int> dictionary, string[] stringArray)
        {
            var partList = new List<PartStruct>();

            foreach (var partLine in stringArray)
            {
                var parts = ConvertFixedLengthString2StringArray(dictionary, partLine);

                // parse the position from (x, y) to x,y
                var position = parts["Position"].Trim(new char[] { '(', ')' });
                var xy = position.Split(' ');

                var part = new PartStruct();
                part.Part = parts["Part"];
                part.Value = parts["Value"];
                part.Package = parts["Package"];
                part.Library = parts["Library"];
                part.Position.X = float.Parse(xy[0]) * convertToMillimeters;
                part.Position.Y = float.Parse(xy[1]) * convertToMillimeters;
                var rotation = parts["Orientation"].TrimStart('M').Replace('R', '-').Replace('L', '+');
                part.Orientation = Convert.ToInt16(rotation);

                partList.Add(part);
            }

            return partList;
        }
        
        static List<PartStruct> ReadPartlist(string filename)
        {
            var partLines = File.ReadAllLines(filename);

            // Remove blank lines
            partLines = partLines.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            partLines = partLines.Skip(4).ToArray();

            var dictionary = ReadLegende(partLines.FirstOrDefault());

            partLines = partLines.Skip(1).ToArray();

            return ConvertStringArrayToPartList(dictionary, partLines);
        }

        static void Main(string[] args)
        {
            // Force decimal point
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB");

            if (args.Length != 2)
            {
                Console.Error.WriteLine($"Not enough arguments, expected 2, got {args.Length}. Exiting");
                return; // expect 2 arguments
            }

            var filenames = new string[] { args[0], args[1] };

            // both files must exist
            for (var i = 0; i < filenames.Length; i++)
            {
                filenames[i] = ResolveShortcut(filenames[i]);

                if (!File.Exists(filenames[i]))
                {
                    Console.Error.WriteLine($"File {filenames[i]} does not exist. Exiting");
                    return;
                }
            }

            var fileLookup = string.Empty;
            var filePartlist = string.Empty;
            foreach (var filename in filenames)
            {
                if (filename.EndsWith(".partlist")) filePartlist = filename;
                if (filename.EndsWith(".lookup")) fileLookup = filename;
            }

            if (string.IsNullOrEmpty(filePartlist) || string.IsNullOrEmpty(fileLookup))
            {
                Console.Error.WriteLine($"Expected a *.partlist and *.lookup file.");
                return;
            }

            //----------------------------------------------------------------------------------------
            // Partlist

            var PartList = ReadPartlist(filePartlist);

            //----------------------------------------------------------------------------------------
            // Lookup
            var PartLookup = new Dictionary<string, string>();

            var partsList = File.ReadAllLines(fileLookup);
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
            var bomFilename = filePartlist + "_bom.csv";

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
            // (Optional) Rotation file
            var rotDictionary = new Dictionary<string, int>();

            var rotFilename = filePartlist + ".rotation";
            if (File.Exists(rotFilename))
            {
                var rotationList = File.ReadAllLines(rotFilename);
                // Remove blank lines
                rotationList = rotationList.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                rotationList = rotationList.Skip(1).ToArray();

                foreach (var line in rotationList)
                {
                    var parts = line.Split(',');
                    rotDictionary.Add(parts[0], Convert.ToInt32(parts[1]));
                }
            }

            //----------------------------------------------------------------------------------------
            // Construct CPL

            var cplFilename = filePartlist + "_cpl.csv";

            sw = new StreamWriter(cplFilename);
            sw.WriteLine("Designator,Mid X,Mid Y,Layer,Rotation");

            foreach (var t in PartList)
            {
                var rotation = t.Orientation;

                // Does this item need additional rotation?
                if (rotDictionary.TryGetValue(t.Part, out var extraRotation))
                    rotation -= extraRotation;

                sw.WriteLine($"{t.Part},{t.Position.X}mm,{t.Position.Y}mm,T,{rotation}");
            }

            sw.Close();

        }
    }
}
