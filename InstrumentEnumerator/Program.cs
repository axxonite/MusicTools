namespace InstrumentEnumerator
{
	#region Usings

	using System.Collections.Generic;
	using System.Diagnostics;
	using System.Diagnostics.CodeAnalysis;
	using System.Drawing;
	using System.IO;
	using System.Linq;
	using System.Text.RegularExpressions;
	using OfficeOpenXml;

	#endregion

	[SuppressMessage("ReSharper", "PossibleNullReferenceException")] sealed class InstrumentEnumerator
	{
		static readonly List<TagDefinition> TagDefinitions = new List<TagDefinition>
			{
				new TagDefinition(Tag.Violin2nd, @"\b(2nd Violins?|Violins? (II|2)|Vlns II)\b"),
				new TagDefinition(Tag.Violin1st, @"\b(1st Violins?|Violins?( (I|1))?|Vln?s)\b"),
				new TagDefinition(Tag.Viola, @"\b(Violas?|Vlas?)\b"),
				new TagDefinition(Tag.Cello, @"\b(Celli|Cellos?)\b"),
				new TagDefinition(Tag.Contrabass, @"\bBass(es)?\b")
			};
		readonly List<string> _instrumentEntries = new List<string>();
		readonly Dictionary<string, Library> _libraries = new Dictionary<string, Library>();
		readonly List<string> _paths = new List<string>();
		ExcelPackage _excel;
		int _instrumentCount;

		public void Enumerate(string spreadsheetFilename, IEnumerable<string> paths, bool organizeInColumns = true)
		{
			var backupFile = Path.GetFileNameWithoutExtension(spreadsheetFilename) + " (Backup).xlsx";
			File.Delete(backupFile);
			if (File.Exists(spreadsheetFilename))
				File.Move(spreadsheetFilename, backupFile);
			var fileInfo = new FileInfo(spreadsheetFilename);
			_excel = new ExcelPackage(fileInfo);

			foreach (var path in paths)
				EnumerateInstruments(path);
			_paths.Add("");
			_paths.Add($"Found {_paths.Count(p => p != "")} unique paths.");
			_instrumentEntries.Add("");
			_instrumentEntries.Add($"Found {_instrumentCount} instrument patches.");
			File.WriteAllLines("Instruments.txt", _instrumentEntries);
			File.WriteAllLines("Paths.txt", _paths);

			foreach (var library in _libraries.Select(kvp => kvp.Value))
			{
				var worksheet = library.Worksheet;
				//var headerFont = worksheet.Cells[1, 1, 1000, 1].Style.Font;
				//headerFont.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));
				worksheet.Cells[1, 1].Value = library.Name;
				worksheet.Cells[1, 1].Style.Font.SetFromFont(new Font("Calibri", 11, FontStyle.Bold));
				var rowPerColumn = Enumerable.Repeat(3, 7).ToArray();
				var folderPerColumn = Enumerable.Repeat("", 7).ToArray();
				foreach (var folder in library.Folders)
				{
					foreach (var instrument in folder.Value.Instruments)
					{
						var entryColumn = GetColumForString(instrument.Name);
						var folderColumn = GetColumForString(folder.Key);
						if (entryColumn == 6 || folderColumn != 6)
							entryColumn = folderColumn;
						if (!organizeInColumns)
							entryColumn = 1;
						if (folderPerColumn[entryColumn] != folder.Key)
						{
							if (rowPerColumn[entryColumn] > 3)
								rowPerColumn[entryColumn]++;
							worksheet.Cells[rowPerColumn[entryColumn], entryColumn].Value = folder.Key;
							worksheet.Cells[rowPerColumn[entryColumn], entryColumn].Style.Font.SetFromFont(new Font("Calibri", 11, FontStyle.Bold));
							rowPerColumn[entryColumn]++;
							folderPerColumn[entryColumn] = folder.Key;
						}
						worksheet.Cells[rowPerColumn[entryColumn]++, entryColumn].Value = instrument.Name;
					}
				}

				worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
			}

			_excel.Save();

			Process.Start(spreadsheetFilename);
		}

		void EnumerateInstruments(string root)
		{
			var files = Directory.EnumerateFiles(root, "*.nki", SearchOption.AllDirectories).ToArray();
			var lastPath = "";
			var regex = new Regex(@"^([\d.]* )");
			LibraryFolder folder = null;
			foreach (var file in files)
			{
				var path = Path.GetDirectoryName(file);
				path = path.Replace(root + @"\", "");
				if (path.Contains("Instruments_old"))
					continue;
				if (path.Contains(@"OT Berlin Strings\Instruments\1.0\"))
					continue;
				path = path.Replace(@"\Instruments\", @"\");
				path = path.Replace(@"BST - Main Collection 2.1\", @"");
				path = path.Replace(@"Instruments_1_2\", @"");
				path = path.Replace(@"Instruments_1_2", @"\");
				path = path.Replace(@"BST A - Special Bows I 2.1\", @"");
				path = path.Replace(@"BST B - Special Bows II 2.1\", @"");
				path = path.Replace(@"BST E - SFX 1.1\", @"");
				//path = path.Replace(@"Instruments main mics\", @"");
				//path = path.Replace(@"Instruments alt mics 1.1\", @"");
				if (lastPath != path)
				{
					var libraryName = path.Split('\\')[0];
					var shortenedLibraryName = libraryName.Length > 30
						                           ? libraryName.Substring(0, 30)
						                           : libraryName;
					Library library;
					if (!_libraries.TryGetValue(shortenedLibraryName, out library))
						library = _libraries[shortenedLibraryName] = new Library(_excel, libraryName);

					path = path.Replace(libraryName + @"\", "");
					folder = library.GetFolder(path);

					if (_instrumentEntries.Any())
						_instrumentEntries.Add("");
					lastPath = path;
					_instrumentEntries.Add(path);
					_paths.Add(path);
				}
				var name = Path.GetFileNameWithoutExtension(file).Replace('_', ' ');
				name = regex.Replace(name, "");
				_instrumentEntries.Add("\t" + name);
				var tags = FindMatchingTags(name);
				var instrument = new Instrument(name, tags);
				folder.Instruments.Add(instrument);
				_instrumentCount++;
			}
		}

		static List<Tag> FindMatchingTags(string s)
		{
			var tags = TagDefinitions.Where(t => t.Matches(s)).Select(t => t.Tag).ToList();
			return tags;
		}

		static int GetColumForString(string s)
		{
			var tags = FindMatchingTags(s);
			if (tags.Any(t => t == Tag.Violin2nd))
				return 2;
			if (tags.Any(t => t == Tag.Violin1st))
				return 1;
			if (tags.Any(t => t == Tag.Viola))
				return 3;
			if (tags.Any(t => t == Tag.Cello))
				return 4;
			return tags.Any(t => t == Tag.Contrabass)
				       ? 5
				       : 6;
		}

		sealed class Library
		{
			public Library(ExcelPackage excel, string name)
			{
				Worksheet = excel.Workbook.Worksheets.Add(name);
				Name = name;
			}

			public Dictionary<string, LibraryFolder> Folders { get; } = new Dictionary<string, LibraryFolder>();
			public string Name { get; }
			public ExcelWorksheet Worksheet { get; }

			public LibraryFolder GetFolder(string path)
			{
				LibraryFolder result;
				if (!Folders.TryGetValue(path, out result))
					result = Folders[path] = new LibraryFolder();
				return result;
			}
		}

		sealed class LibraryFolder
		{
			public List<Instrument> Instruments { get; } = new List<Instrument>();
		}

		enum Tag
		{
			// ReSharper disable InconsistentNaming
			Violin1st,
			Violin2nd,
			Viola,
			Cello,
			Contrabass
			// ReSharper restore InconsistentNaming
		}

		struct Instrument
		{
			public Instrument(string name, List<Tag> tags)
			{
				Name = name;
				Tags = tags;
			}

			public string Name { get; }
			public List<Tag> Tags { get; }
		}

		struct TagDefinition
		{
			readonly Regex _regex;

			public TagDefinition(Tag tag, string regex)
			{
				Tag = tag;
				_regex = new Regex(regex);
			}

			public Tag Tag { get; }

			public bool Matches(string s) => _regex.IsMatch(s);
		}
	}

	static class Program
	{
		static void Main()
		{
			var enumerator = new InstrumentEnumerator();
			enumerator.Enumerate("String Libraries.xlsx", new[] { @"E:\Sample Libraries\Kontakt\Strings\Ensemble", @"F:\Sample Libraries\Kontakt\Strings\Ensemble" });
			enumerator = new InstrumentEnumerator();
			enumerator.Enumerate("Orchestra Libraries.xlsx", new[] { @"E:\Sample Libraries\Kontakt\Orchestras", @"F:\Sample Libraries\Kontakt\Orchestras" }, false);
		}
	}
}