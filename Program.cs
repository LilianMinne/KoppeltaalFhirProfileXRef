using Hl7.Fhir.Model;
using Hl7.Fhir.Specification;
using Hl7.Fhir.Specification.Snapshot;
using Hl7.Fhir.Specification.Source;
using Hl7.Fhir.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace FhirProfileXRef
{
    class Program
    {
        const bool includeSubDirs = true;

        static readonly string TYPENAME_REFERENCE = FHIRAllTypes.Reference.GetLiteral();

        static string _targetPath;

        static ZipSource _coreSource;
        static CachedResolver _cachedCoreSource;

        static DirectorySource _dirSource;
        static CachedResolver _cachedDirSource;
        static SnapshotGenerator _snapshotGenerator;

        static List<string> _coreProfiles;
        static List<string> _userProfiles;

        static Dictionary<string, string> _mappings = new Dictionary<string, string>();

        static Application xlApp = new Application();
        static Workbook wb;
        static Worksheet ws;
        static int count = 2;

        static void Main(string[] args)
        {
            ShowIntro();

            if (InitArguments(args))
            {
                Console.WriteLine($"Location: '{_targetPath}'");
                InitCoreProfiles();
                InitUserProfiles();
                CreateExcelSheet();
                ValidateXRef();
                SaveExcelSheet();
            }
            else
            {
                ShowHelp();
            }

//#if DEBUG
            Console.ReadLine();
//#endif
        }

        static void CreateExcelSheet()
        {
            wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            ws = (Worksheet)wb.Worksheets[1];

            ws.Columns[1].ColumnWidth = 20;
            ws.Columns[2].ColumnWidth = 30;
            ws.Columns[3].ColumnWidth = 55;
            ws.Columns[4].ColumnWidth = 55;

            ws.Rows[1].Font.Bold = true;

            ws.Cells[1, 1] = "Resource name";
            ws.Cells[1, 2] = "Element path";
            ws.Cells[1, 3] = "Reference found";
            ws.Cells[1, 4] = "Reference suggestion";
        }

        static void SaveExcelSheet()
        {
            if (System.IO.File.Exists(@"C:\Git\KoppeltaalFhirProfileXRef\FhirProfileXRef.xlsx"))
            {
                System.IO.File.Delete(@"C:\Git\KoppeltaalFhirProfileXRef\FhirProfileXRef.xlsx");
            }

            wb.SaveAs(@"C:\Git\KoppeltaalFhirProfileXRef\FhirProfileXRef.xlsx");
            xlApp.Workbooks.Close();
        }

        static bool InitArguments(string[] args)
        {
            if (args.Length == 0)
            {
                _targetPath = Directory.GetCurrentDirectory();
                return true;
            }
            if (args.Length == 1)
            {
                var path = args[0];
                if (Directory.Exists(path))
                {
                    _targetPath = args[0];
                    return true;
                }
            }
            return false;
        }

        static void ShowIntro()
        {
            var title = GetAppTitle();
            Console.WriteLine(title);
            Console.WriteLine(new string('=', title.Length));
            Console.WriteLine("FHIR profile cross references validator");
            Console.WriteLine("(C) Furore 2017");
            Console.WriteLine();
        }

        static void ShowHelp()
        {
            Console.WriteLine("Usage: ");
            Console.WriteLine($"{ExeTitle} [path]");
        }

        static void InitCoreProfiles()
        {

            Console.WriteLine("Load FHIR core resource definitions...");

            var src = _coreSource = ZipSource.CreateValidationSource();
            _cachedCoreSource = new CachedResolver(src);
            var profiles = _coreProfiles = src.ListResourceUris(ResourceType.StructureDefinition).ToList();

            Console.WriteLine($"Found {profiles.Count} core definitions.");
        }

        static void InitUserProfiles()
        {
            Console.WriteLine($"Fetch profiles in target location...");
            var src = _dirSource = new DirectorySource(_targetPath, includeSubDirs);
            _cachedDirSource = new CachedResolver(src);

            var userProfiles = _userProfiles = src.ListResourceUris().ToList();
            Console.WriteLine($"Found {userProfiles.Count} profiles.");

            Console.WriteLine($"Determine mappings...");
            foreach (var profile in userProfiles)
            {
                var sd = _cachedDirSource.FindStructureDefinition(profile, false);
                if (EnsureSnapshot(sd))
                {
                    var key = ModelInfo.CanonicalUriForFhirCoreType(sd.Type);
                    if (!_mappings.TryGetValue(key, out string existing))
                    {
                        Console.WriteLine($"Map references of type '{sd.Type}' to user profile '{sd.Url}'");
                        _mappings.Add(key, sd.Url);
                    }
                    else
                    {
                        Console.WriteLine($"Warning! Ignore duplicate user profile '{sd.Url}' for reference target type '{sd.Type}'");
                    }
                }
            }
        }

        static void ValidateXRef()
        {
            Console.WriteLine($"Validate x-refs...");
            foreach (var profile in _userProfiles)
            {
                Validate(profile);
            }
        }

        static void Validate(string profileUrl)
        {
            Console.WriteLine($"Validate '{profileUrl}' ...");
            var sd = _cachedDirSource.FindStructureDefinition(profileUrl, false);

            if (EnsureSnapshot(sd))
            {
                foreach (var elem in sd.Snapshot.Element)
                {
                    ValidateElement(elem, sd);
                }
            }
        }

        static void ValidateElement(ElementDefinition elem, StructureDefinition sd)
        {
            foreach (var type in elem.Type)
            {
                if (type.Code == TYPENAME_REFERENCE)
                {
                    var tgt = type.TargetProfile;
                    // Console.WriteLine($"'{elem.Path}' => '{tgt}'");
                    if (_mappings.TryGetValue(tgt, out string profile))
                    {
                        Console.WriteLine($"Warning! '{elem.Path}' : '{tgt}' => '{profile}'");
                        ws.Cells[count, 1] = sd.Name;
                        ws.Cells[count, 2] = elem.Path;
                        ws.Cells[count, 3] = tgt;
                        ws.Cells[count, 4] = profile;
                        count++;
                    }
                }
            }
        }

        static bool EnsureSnapshot(StructureDefinition sd)
        {
            if (!sd.HasSnapshot)
            {
                var generator = GetSnapshotGenerator();
                Console.WriteLine($"Generate snapshot for profile '{sd.Url}' ...");
                generator.Update(sd);
                if (generator.Outcome != null)
                {
                    Console.WriteLine("Snapshot generator returned one or more issues:");
                    foreach (var issue in generator.Outcome.Issue)
                    {
                        Console.WriteLine($"{issue.Details}");
                    }
                    return false;
                }
            }
            return true;
        }

        static SnapshotGenerator GetSnapshotGenerator()
        {
            if (_snapshotGenerator == null)
            {
                var src = new MultiResolver(_cachedCoreSource, _cachedDirSource);
                _snapshotGenerator = new SnapshotGenerator(src);
            }
            return _snapshotGenerator;
        }

        static string GetAppTitle()
        {
            AssemblyTitleAttribute titleAttr = Attribute.GetCustomAttribute(
                Assembly.GetExecutingAssembly(),
                typeof(AssemblyTitleAttribute),
                false) as AssemblyTitleAttribute;
            return titleAttr?.Title;
        }

        static string ExeTitle => Path.GetFileName(Assembly.GetEntryAssembly().CodeBase);

    }
}
