namespace ExpressionMapBuilder
{
	#region Usings

	using System;
	using System.Collections.Generic;
	using System.Diagnostics;
	using System.Diagnostics.CodeAnalysis;
	using System.IO;
	using System.Linq;
	using System.Text.RegularExpressions;
	using HtmlAgilityPack;
	using MoreLinq;

	#endregion

	struct Articulation
	{
		public Articulation(string name, int outputKS, int channelOffset, bool keyswitched, int inputKS)
		{
			Name = name.TrimEnd('*');
			ExcludeFromRemote = name.EndsWith("*", StringComparison.Ordinal);
			ChannelOffset = channelOffset;
			Keyswitched = keyswitched;
			OutputKS = outputKS;
			InputKS = inputKS;
		}

		public int ChannelOffset { get; }
		public bool ExcludeFromRemote { get; }
		public int OutputKS { get; }
		public bool Keyswitched { get; }
		public string Name { get; }
		public int InputKS { get; }
	}

	sealed class ExpressionMap
	{
		public ExpressionMap(string name, int rootKey)
		{
			Name = name;
			RootKey = rootKey;
		}

		public ExpressionMap(string name, ExpressionMap baseMap)
		{
			Name = name;
			RootKey = baseMap.RootKey;
			Articulations = new List<Articulation>(baseMap.Articulations);
			BaseChannel = baseMap.BaseChannel;
			Remotes = new List<Articulation>(baseMap.Remotes);
			StandardRemoteAssignments = baseMap.StandardRemoteAssignments;
			InheritedArticulations = true;
		}

		public List<Articulation> Articulations { get; private set; } = new List<Articulation>();
		public int BaseChannel { get; set; } = -1;
		public string Name { get; }
		public List<Articulation> Remotes { get; private set; } = new List<Articulation>();
		public int RootKey { get; private set; }
		public bool StandardRemoteAssignments { get; set; } = true;
		public bool InheritedArticulations { get; set; }

		public void AssignRemotes()
		{
			if (!Remotes.Any())
				Remotes.AddRange(Articulations.Where(a => !a.ExcludeFromRemote));
		}

		public void ChangeRootKey(int newKey)
		{
			var diff = newKey - RootKey;
			Articulations = Articulations.Select(a => new Articulation(a.Name, a.OutputKS + diff, a.ChannelOffset, a.Keyswitched, a.InputKS)).ToList();
			Remotes = Remotes.Select(a => new Articulation(a.Name, a.OutputKS + diff, a.ChannelOffset, a.Keyswitched, a.InputKS)).ToList();
			RootKey = newKey;
		}
	}

	sealed class ExpressionMapBuilder
	{
		static readonly string[] NoteNames =
			{
				"C-2", "C#-2", "D-2", "D#-2", "E-2", "F-2", "F#-2", "G-2", "G#-2", "A-2", "A#-2", "B-2", "C-1", "C#-1", "D-1", "D#-1", "E-1", "F-1", "F#-1", "G-1", "G#-1", "A-1", "A#-1",
				"B-1", "C0", "C#0", "D0", "D#0", "E0", "F0", "F#0", "G0", "G#0", "A0", "A#0", "B0", "C1", "C#1", "D1", "D#1", "E1", "F1", "F#1", "G1", "G#1", "A1", "A#1", "B1", "C2", "C#2", "D2", "D#2", "E2", "F2", "F#2",
				"G2", "G#2", "A2", "A#2", "B2", "C3", "C#3", "D3", "D#3", "E3", "F3", "F#3", "G3", "G#3", "A3", "A#3", "B3", "C4", "C#4", "D4", "D#4", "E4", "F4", "F#4", "G4", "G#4", "A4", "A#4", "B4", "C5", "C#5", "D5",
				"D#5", "E5", "F5", "F#5", "G5", "G#5", "A5", "A#5", "B5", "C6", "C#6", "D6", "D#6", "E6", "F6", "F#6", "G6", "G#6", "A6", "A#6", "B6", "C7", "C#7", "D7", "D#7", "E7", "F7", "F#7", "G7", "G#7", "A7", "A#7",
				"B7", "C8", "C#8", "D8", "D#8", "E8", "F8", "F#8", "G8"
			};

		readonly Tuple<string, int, bool>[] _inputKSNames =
			{
				Tuple.Create("leg", 0, true), Tuple.Create("long", 1, true), Tuple.Create("sus", 1, false), Tuple.Create("short", 2, true), Tuple.Create("stac", 2, false), Tuple.Create("spic", 3, false),
				Tuple.Create("pizz", 4, false), Tuple.Create("trem", 5, false), Tuple.Create("trills+1", 6, false), Tuple.Create("trills+2", 7, false), Tuple.Create("trills", 6, true)
			};

		Dictionary<string, ExpressionMap> ExpressionMaps { get; } = new Dictionary<string, ExpressionMap>();

		public void BuildExpressionMaps(string filename)
		{
			ParseMapFile(filename);
			SaveExpressionMaps();
		}

		static HtmlDocument CreateHtmlDocument(string contents)
		{
			var document = new HtmlDocument();
			document.LoadHtml(contents);
			return document;
		}

		// todo I want to be able to define the various 8dio instruments as separate entries for their articulations, and then load these separate entries in a different expression map and refer to them there.
		// this is because I plan to have 8dio function in one of two modes: either a pass through mode with 3 separate tracks for the different articulations available, or a combined mode with a select set of
		// articulations grouped up under one track.
		// why do I need separate entries again? Can't I just refer to the existing entries? I think so...

		[SuppressMessage("ReSharper", "PossibleNullReferenceException")] void ParseMapFile(string filename)
		{
			var articulationsRegex = new Regex(@" *([\w\-+/*() ]+)(?:\[(?:(Ch|NoKS|InputKS)(\+?\d+)?\|?)+\])?");
			var expressionMapRegEx = new Regex(@"(Map|Base|RootKey|BaseChannel|Art|Remotes|StdRemotes)\s*(\[[^]]*\])?:\s*(?:([^,]+),?)+");
			var defaultRootKey = Array.IndexOf(NoteNames, "C0");

			var lines = File.ReadAllLines(filename).Where(s => s != "");
			ExpressionMap map = null;
			var key = defaultRootKey;
			foreach (var line in lines)
			{
				var match = expressionMapRegEx.Matches(line)[0];
				Debug.Assert(match.Groups[1].Captures[0].Value.ToLower() == "map" || map != null, "Must declare map block first.");
				switch (match.Groups[1].Captures[0].Value.ToLower())
				{
					case "map":
						map?.AssignRemotes();
						var mapName = match.Groups[3].Captures[0].Value;
						ExpressionMaps[mapName] = map = new ExpressionMap(mapName, defaultRootKey);
						key = defaultRootKey;
						break;
					case "base":
						var baseMap = ExpressionMaps[match.Groups[3].Captures[0].Value];
						map = ExpressionMaps[map.Name] = new ExpressionMap(map.Name, baseMap);
						key = map.RootKey;
						break;
					case "rootkey":
						var newKey = Array.IndexOf(NoteNames, match.Groups[3].Captures[0].Value);
						map.ChangeRootKey(newKey);
						key = newKey;
						break;
					case "basechannel":
						map.BaseChannel = int.Parse(match.Groups[3].Captures[0].Value) - 1;
						break;
					case "stdremotes":
						map.StandardRemoteAssignments = match.Groups[3].Captures[0].Value.ToLower() == "true";
						break;
					case "art":
						if ( map.InheritedArticulations)
						{
							map.Articulations.Clear();
							map.InheritedArticulations = false;
						}
                        key = defaultRootKey;
						map.Remotes.Clear();
						var defaultAttributesText = match.Groups[2].Captures.Count > 0  ? ("Default" + match.Groups[2].Captures[0].Value) : "";
						foreach (Capture articulationEntry in match.Groups[3].Captures)
						{
						    var articulation = new Articulation("", -1, -1, true, -1);
						    if (defaultAttributesText != "")
						        articulation = ParseArticulationAttributes(articulation, defaultAttributesText, key);
						    articulation = ParseArticulationAttributes(articulation, articulationEntry.Value, key);
                            if (articulation.Keyswitched)
								key++;
							if (articulation.Name.ToLower() != "none")
								map.Articulations.Add(articulation);
						}
						break;
					case "remotes":
						map.Remotes.Clear();
						foreach (Capture remoteEntry in match.Groups[3].Captures)
						{
							var remoteName = remoteEntry.Value.ToLower();
						    if (remoteName.ToLower() == "none")
						        map.Remotes.Add(new Articulation("None", -1, -1, false, -1));
						    else
						    {
						        var articulation = map.Articulations.First(c => c.Name.ToLower().StartsWith(remoteName, StringComparison.Ordinal));
						        map.Remotes.Add(articulation);
						    }
                        }
						break;
					default:
						Debug.Assert(false);
						break;
				}
			}
			map?.AssignRemotes();
		}

	    Articulation ParseArticulationAttributes(Articulation articulation, string declaration, int key)
	    {
	        var channelOffset = articulation.ChannelOffset;
	        var inputKS = articulation.InputKS;
	        var keyswitched = articulation.Keyswitched;
            var articulationsRegex = new Regex(@" *([\w\-+/*() ]+)(?:\[(?:(Ch|NoKS|InputKS)(\+?\d+)?\|?)+\])?");
            var articulationMatch = articulationsRegex.Matches(declaration)[0];
	        var attributes = articulationMatch.Groups[2].Captures.Cast<Capture>().ToList();
	        foreach (Capture attribute in attributes)
	        {
	            switch (attribute.Value.ToLower())
	            {
	                case "noks":
	                    keyswitched = false;
	                    break;
	                case "ch":
	                    channelOffset = int.Parse(articulationMatch.Groups[3].Captures[0].Value);
	                    break;
	                case "inputks":
	                    inputKS = int.Parse(articulationMatch.Groups[3].Captures[0].Value) - 1;
	                    break;
	                default:
	                    Debug.Assert(false);
	                    break;
	            }
	        }
	        return new Articulation(articulationMatch.Groups[1].Captures[0].Value, key, channelOffset, keyswitched, inputKS);
	    }

		void SaveExpressionMaps()
		{
			var templateFile = File.ReadAllText("Template.expressionmap");
			foreach (var map in ExpressionMaps.Select(kvp => kvp.Value))
			{
				var document = CreateHtmlDocument(templateFile);
				var rootNode = document.DocumentNode;
				var nextAvailableRemote = map.StandardRemoteAssignments
					                          ? 8
					                          : 0;
				var inputKSSlots = rootNode.SelectNodes("//obj[@class='PSoundSlot']").ToArray();
				var mapNameNode = rootNode.SelectSingleNode("instrumentmap/string[@name='name']");
				mapNameNode.SetAttributeValue("value", map.Name);

				foreach (var remote in map.Remotes)
				{
				    if (remote.Name == "None")
				        continue;
					var inputKS = remote.InputKS;
					if (inputKS == -1 && map.StandardRemoteAssignments)
						inputKS = _inputKSNames.FirstOrDefault(s => inputKSSlots[s.Item2] != null && (s.Item3 ? remote.Name.ToLower().StartsWith(s.Item1, StringComparison.Ordinal) : remote.Name.ToLower().Contains(s.Item1)))?.Item2 ?? -1;
					if (inputKS == -1)
						inputKS = nextAvailableRemote++;
					var inputKSSlot = inputKSSlots[inputKS];
					inputKSSlot.SelectSingleNode("member[@name='name']/string").SetAttributeValue("value", remote.Name);
					var keyswitchNode = inputKSSlot.SelectSingleNode(".//obj[@class='POutputEvent']");
					if (!remote.Keyswitched)
						keyswitchNode.Remove();
					else
						keyswitchNode.SelectSingleNode("int[@name='data1']").SetAttributeValue("value", remote.OutputKS.ToString());
					if (remote.ChannelOffset != -1)
					{
						var noteChangerNode = inputKSSlot.SelectSingleNode(".//obj[@class='PSlotNoteChanger']/int[@name='channel']");
						noteChangerNode.SetAttributeValue("value", (map.BaseChannel + remote.ChannelOffset).ToString());
					}
					inputKSSlots[inputKS] = null;
				}

				inputKSSlots.Where(s => s != null).ForEach(s => s.Remove());

				if (!map.StandardRemoteAssignments)
					rootNode.SelectNodes("//obj[@class='USlotVisuals']").ForEach(s => s.Remove());
				else
				{
					// Remove unused articulation visual slots.
					var slotVisuals = rootNode.SelectNodes("//member[@name='slotvisuals']//obj[@class='USlotVisuals']").ToArray();
					for (var i = 0; i < slotVisuals.Length; i++)
						if (inputKSSlots[i] != null)
							slotVisuals[i].Remove();
				}

				var output = document.DocumentNode.OuterHtml;
				output = output.Replace("id=", "ID=");
				output = output.Replace("instrumentmap", "InstrumentMap");
				File.WriteAllText($@"Maps\{map.Name}.expressionMap", output);
			}
		}
	}

	static class Program
	{
		static void Main(string[] args)
		{
			var builder = new ExpressionMapBuilder();
			builder.BuildExpressionMaps(args[0]);
		}
	}
}