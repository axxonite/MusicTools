namespace DrumMapConverter
{
	#region Usings

	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using System.Xml;
	using OfficeOpenXml;

	#endregion

	static class Program
	{
		static readonly string[] NoteNames =
			{
				"C-2", "C#-2", "D-2", "D#-2", "E-2", "F-2", "F#-2", "G-2", "G#-2", "A-2", "A#-2", "B-2", "C-1", "C#-1", "D-1", "D#-1", "E-1", "F-1", "F#-1", "G-1", "G#-1", "A-1", "A#-1",
				"B-1", "C0", "C#0", "D0", "D#0", "E0", "F0", "F#0", "G0", "G#0", "A0", "A#0", "B0", "C1", "C#1", "D1", "D#1", "E1", "F1", "F#1", "G1", "G#1", "A1", "A#1", "B1", "C2", "C#2", "D2", "D#2", "E2", "F2", "F#2",
				"G2", "G#2", "A2", "A#2", "B2", "C3", "C#3", "D3", "D#3", "E3", "F3", "F#3", "G3", "G#3", "A3", "A#3", "B3", "C4", "C#4", "D4", "D#4", "E4", "F4", "F#4", "G4", "G#4", "A4", "A#4", "B4", "C5", "C#5", "D5",
				"D#5", "E5", "F5", "F#5", "G5", "G#5", "A5", "A#5", "B5", "C6", "C#6", "D6", "D#6", "E6", "F6", "F#6", "G6", "G#6", "A6", "A#6", "B6", "C7", "C#7", "D7", "D#7", "E7", "F7", "F#7", "G7", "G#7", "A7", "A#7",
				"B7", "C8", "C#8", "D8", "D#8", "E8", "F8", "F#8", "G8"
			};

		static void Main(string[] args)
		{
			var noteIndices = new Dictionary<string, int>();
			for (var i = 0; i < NoteNames.Length; i++)
				noteIndices[NoteNames[i]] = i;
			var files = Directory.EnumerateFiles(args[0], "*.xlsx").ToArray();
			foreach (var file in files)
			{
				var fileInfo = new FileInfo(file);
				var excel = new ExcelPackage(fileInfo);
				var worksheet = excel.Workbook.Worksheets["Drum Map"];

				var dictIndex = 0;
				var drumMap = Enumerable.Repeat("", 128).ToDictionary(i => dictIndex++);
				var hasMultiples = new HashSet<string>();
				var nameCounts = new Dictionary<string, int>();
				for (var row = 2; row <= worksheet.Dimension.End.Row; row++)
				{
					var index = noteIndices[worksheet.Cells[row, 1].Text];
					var name = drumMap[index] = worksheet.Cells[row, 2].Text;
					if (nameCounts.ContainsKey(name))
						hasMultiples.Add(name);
					nameCounts[name] = 1;
				}
				for (var row = 2; row <= worksheet.Dimension.End.Row; row++)
				{
					worksheet.Cells[row, 1].Clear();
					worksheet.Cells[row, 2].Clear();
				}
				for (var i = 0; i < drumMap.Count; i++)
				{
					worksheet.Cells[i + 2, 1].Value = NoteNames[i];
					worksheet.Cells[i + 2, 2].Value = drumMap[i];
				}
				excel.Save();

				for (var i = 0; i < drumMap.Count; i++)
				{
					var name = drumMap[i];
					if (name != "" && hasMultiples.Contains(name))
					{
						var nameCount = nameCounts[name];
						nameCounts[name] = nameCount + 1;
						drumMap[i] = name + $" {nameCount}";
					}
				}

				var sortedDrumMap = drumMap.OrderBy(
					kvp => kvp.Value != ""
						       ? 0
						       : 1).ThenByDescending(
					kvp => kvp.Value != ""
						       ? kvp.Key
						       : -kvp.Key).Select(kvp => new { Index = kvp.Key.ToString(), Name = kvp.Value }).ToList();
				var ws = new XmlWriterSettings { Indent = true, NewLineOnAttributes = false };
				using (var writer = XmlWriter.Create(Path.ChangeExtension(file, "drm"), ws))
				{
					writer.WriteStartElement("DrumMap");
					WriteElement(writer, "string", "Name", Path.GetFileNameWithoutExtension(file), "wide", "true");
					writer.WriteStartElement("list");
					writer.WriteAttributeString("name", "Quantize");
					writer.WriteAttributeString("type", "list");
					writer.WriteStartElement("item");
					WriteElement(writer, "int", "Grid", "4");
					WriteElement(writer, "int", "Type", "0");
					WriteElement(writer, "float", "Swing", "0");
					WriteElement(writer, "int", "Legato", "50");
					writer.WriteEndElement();
					writer.WriteEndElement();

					writer.WriteStartElement("list");
					writer.WriteAttributeString("name", "Map");
					writer.WriteAttributeString("type", "list");
					foreach (var entry in drumMap)
					{
						writer.WriteStartElement("item");
						WriteElement(writer, "int", "INote", entry.Key.ToString());
						WriteElement(writer, "int", "ONote", entry.Key.ToString());
						WriteElement(writer, "int", "Channel", "-1");
						WriteElement(writer, "float", "Length", "200");
						WriteElement(writer, "int", "Mute", "0");
						WriteElement(writer, "int", "DisplayNote", entry.Key.ToString());
						WriteElement(writer, "int", "HeadSymbol", "0");
						WriteElement(writer, "int", "Voice", "0");
						WriteElement(writer, "int", "PortIndex", "0");
						WriteElement(writer, "string", "Name", entry.Value, "wide", "true");
						WriteElement(writer, "int", "QuantizeIndex", "0");
						writer.WriteEndElement();
					}
					writer.WriteEndElement();

					writer.WriteStartElement("list");
					writer.WriteAttributeString("name", "Order");
					writer.WriteAttributeString("type", "int");
					foreach (var entry in sortedDrumMap)
					{
						writer.WriteStartElement("item");
						writer.WriteAttributeString("value", entry.Index);
						writer.WriteEndElement();
					}
					writer.WriteEndElement();

					writer.WriteStartElement("list");
					writer.WriteAttributeString("name", "OutputDevices");
					writer.WriteAttributeString("type", "list");
					writer.WriteStartElement("item");
					WriteElement(writer, "string", "DeviceName", "Default Device");
					WriteElement(writer, "string", "PortName", "Default Port");
					writer.WriteEndElement();
					writer.WriteEndElement();
					WriteElement(writer, "int", "Flags", "0");
					writer.WriteEndElement();
				}
			}
		}

		static void WriteElement(XmlWriter writer, string type, string name, string value, string extraAttributeName = "", string extraAttributeValue = "")
		{
			writer.WriteStartElement(type);
			writer.WriteAttributeString("name", name);
			writer.WriteAttributeString("value", value);
			if (extraAttributeName != "")
				writer.WriteAttributeString(extraAttributeName, extraAttributeValue);
			writer.WriteEndElement();
		}
	}
}