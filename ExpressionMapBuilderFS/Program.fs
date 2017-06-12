open System
open System.IO
open System.Diagnostics
open System.Text.RegularExpressions
open HtmlAgilityPack

type Articulation = 
    { name : string
      excludeFromRemote : bool
      channelOffset: int
      key : int
      keyswitched : bool
      remoteSlot : int }
let defaultArticulation = {name = ""; channelOffset = 0; key = 0; keyswitched = true; remoteSlot = 0; excludeFromRemote = false}

type ExpressionMap = 
    { articulations : List<Articulation>
      baseChannel : int
      name : string
      remotes : List<Articulation>
      rootKey : int
      standardRemoteAssignments : bool }
let defaultExpressionMap = { articulations = List.Empty; baseChannel = -1; name = ""; remotes = List.Empty; rootKey = 0; standardRemoteAssignments = true }
 
let noteNames = [|"C-2"; "C#-2"; "D-2"; "D#-2"; "E-2"; "F-2"; "F#-2"; "G-2"; "G#-2"; "A-2"; "A#-2"; "B-2"; "C-1"; "C#-1"; "D-1"; "D#-1"; "E-1"; "F-1"; "F#-1"; "G-1"; "G#-1"; "A-1"; "A#-1";
                  "B-1"; "C0"; "C#0"; "D0"; "D#0"; "E0"; "F0"; "F#0"; "G0"; "G#0"; "A0"; "A#0"; "B0"; "C1"; "C#1"; "D1"; "D#1"; "E1"; "F1"; "F#1"; "G1"; "G#1"; "A1"; "A#1"; "B1"; "C2"; "C#2"; "D2"; "D#2"; "E2"; "F2"; "F#2";
                  "G2"; "G#2"; "A2"; "A#2"; "B2"; "C3"; "C#3"; "D3"; "D#3"; "E3"; "F3"; "F#3"; "G3"; "G#3"; "A3"; "A#3"; "B3"; "C4"; "C#4"; "D4"; "D#4"; "E4"; "F4"; "F#4"; "G4"; "G#4"; "A4"; "A#4"; "B4"; "C5"; "C#5"; "D5";
                  "D#5"; "E5"; "F5"; "F#5"; "G5"; "G#5"; "A5"; "A#5"; "B5"; "C6"; "C#6"; "D6"; "D#6"; "E6"; "F6"; "F#6"; "G6"; "G#6"; "A6"; "A#6"; "B6"; "C7"; "C#7"; "D7"; "D#7"; "E7"; "F7"; "F#7"; "G7"; "G#7"; "A7"; "A#7";
                  "B7"; "C8"; "C#8"; "D8"; "D#8"; "E8"; "F8"; "F#8"; "G8" |]

let assignRemotes (m:ExpressionMap) =
    if (not m.remotes.IsEmpty) then
        {m with remotes = List.append m.remotes (List.filter(fun x -> not x.excludeFromRemote) m.articulations) }
    else m

let changeRootKey (m:ExpressionMap) newKey =
    let diff = newKey - m.rootKey
    {m with articulations = List.map (fun a -> { a with key = a.key + diff }) m.articulations; remotes = List.map (fun r -> { r with key = r.key + diff }) m.remotes; rootKey = newKey }

let parseMapFile f = 
    let articulationsRegex = new Regex(@" *([\w\-+/*() ]+)(?:\[(?:(Ch|NoKS|Slot)(\+?\d+)?\|?)+\])?")
    let expressionMapRegex = new Regex(@"(Map|Base|RootKey|BaseChannel|Art|Remote|StdRemotes)\s*:\s*(?:([^,]+),?)+")
    let defaultRootKey = Array.findIndex(fun x -> x = "C0") noteNames
    let lines = File.ReadAllLines(f) |> Seq.where(fun x -> x <> "")
    let mutable expressionMaps = Map.empty<string, ExpressionMap>
    let mutable map = defaultExpressionMap
    let mutable key = defaultRootKey
    for line in lines do
        let lineMatch = expressionMapRegex.Match(line)
        Debug.Assert(lineMatch.Groups.[1].Captures.[0].Value.ToLower() = "map" || map.name <> "", "Must declare map block first.")
        match lineMatch.Groups.[1].Captures.[0].Value.ToLower() with
        | "map" ->
            if ( map.name <> "" ) then
                map <- assignRemotes map
                expressionMaps <- expressionMaps |> Map.add map.name map
            map <- {defaultExpressionMap with name = lineMatch.Groups.[2].Captures.[0].Value; rootKey = defaultRootKey }
            key <- defaultRootKey
        | "base" ->
            if ( map.name <> "" ) then
                expressionMaps <- expressionMaps |> Map.add map.name map
            let baseMap = expressionMaps |> Map.find lineMatch.Groups.[2].Captures.[0].Value
            map <- {baseMap with name = map.name; rootKey = map.rootKey }
            key <- defaultRootKey
        | "rootkey" ->
            let newKey =  Array.findIndex(fun x -> x = lineMatch.Groups.[2].Captures.[0].Value) noteNames
            map <- changeRootKey map newKey     
            key <- newKey
        | "basechannel" ->
            map <- {map with baseChannel = (lineMatch.Groups.[2].Captures.[0].Value |> int) - 1}
        | "stdremotes" ->
            map <- {map with standardRemoteAssignments = lineMatch.Groups.[2].Captures.[0].Value.ToLower() <> "true"}
        | "art" ->
            let parseArticulation (entry:Capture) =
                let articulationMatch = articulationsRegex.Matches(entry.Value).[0]
                let mutable articulation = { defaultArticulation with name = articulationMatch.Groups.[1].Captures.[0].Value; key = key }
                for attribute in articulationMatch.Groups.[2].Captures do
                    match attribute.Value.ToLower() with
                        | "noks" -> articulation <- { articulation with keyswitched = false }
                        | "ch" -> articulation <- { articulation with channelOffset = articulationMatch.Groups.[3].Captures.[0].Value |> int }
                        | "slot" -> articulation <- { articulation with remoteSlot =  (articulationMatch.Groups.[3].Captures.[0].Value |> int) - 1 }
                        | _ -> Debug.Assert(false)
                if (articulation.keyswitched) then
                    key <- key + 1
                (articulation)
            let articulations = lineMatch.Groups.[2].Captures |> Seq.cast<Capture> |> Seq.map ( fun a -> parseArticulation a ) |> Seq.filter ( fun a -> a.name.ToLower() <> "none" ) |> Seq.toList
            map <- { map with articulations = map.articulations |> List.append articulations; remotes = List.empty }
        | "remote" ->
            let mapRemote (remoteEntry:Capture) =
                let remoteName = remoteEntry.Value.ToLower()
                map.articulations |> List.find (fun a -> a.name.ToLower().StartsWith(remoteName))
            let remotes = lineMatch.Groups.[2].Captures |> Seq.cast<Capture> |> Seq.map (fun a -> mapRemote a) |> Seq.toList
            map <- { map with remotes = remotes }
        | _ -> Debug.Assert(false)
    if ( map.name <> "" ) then
        map <- assignRemotes map
        expressionMaps <- expressionMaps |> Map.add map.name map
    (expressionMaps)

let createHtmlDocument contents =
    let document = new HtmlDocument()
    document.LoadHtml(contents)
    (document)

let soundSlotNames = [| ("leg", 0, true); ("long", 1, true); ("sus", 1, false); ("short", 2, true); ("stac", 2, false); ("spic", 3, false); ("pizz", 4, false); ("trem", 5, false); ("trills+1", 6, false); ("trills+2", 7, false); ("trills", 6, true) |]

let saveExpressionMaps (maps : Map<string, ExpressionMap>) =
    let templateFile = File.ReadAllText("Template.expressionmap")
    for map in maps |> Map.toList |> List.map(fun (_,m) -> m) do
        let document = createHtmlDocument templateFile
        let rootNode = document.DocumentNode
        let mutable nextAvailableRemote = 8
        if (not map.standardRemoteAssignments) then
            nextAvailableRemote <- 0
        let soundSlots = rootNode.SelectNodes("//obj[@class='PSoundSlot']") |> Seq.cast<HtmlNode option> |> Seq.toArray
        let mapNameNode = rootNode.SelectSingleNode("instrumentmap/string[@name='name']")
        for remote in map.remotes do
            let assignNextRemote () =
                let slot = nextAvailableRemote
                nextAvailableRemote <- nextAvailableRemote + 1
                slot
            let matchSlot () =
                soundSlotNames |> Array.tryFind(fun (name,index, mustStartWith) -> 
                    soundSlots.[index] <> None && 
                    match mustStartWith with
                    | true -> remote.name.ToLower().StartsWith(name)
                    | false -> remote.name.ToLower().Contains(name))
            let remoteSlot = 
                match remote.remoteSlot with
                | -1 ->
                    match map.standardRemoteAssignments with
                    | true -> 
                        match matchSlot() with
                        | Some slot -> 
                            let _,index,_ = slot
                            index
                        | None -> assignNextRemote()
                    | false -> assignNextRemote()
                | _ -> remote.remoteSlot
            let soundSlot = soundSlots.[remoteSlot].Value
            soundSlot.SelectSingleNode("member[@name='name']/string").SetAttributeValue("value", remote.name) |> ignore
            let keyswitchNode = soundSlot.SelectSingleNode(".//obj[@class='POutputEvent']");
            match remote.keyswitched with
            | false -> keyswitchNode.Remove()
            | true -> keyswitchNode.SelectSingleNode("int[@name='data1]").SetAttributeValue("value", remote.key.ToString()) |> ignore

            if (remote.channelOffset <> -1 ) then
                let noteChangerNode = soundSlot.SelectSingleNode(".//obj[@class='PSlotNoteChanger']/int[@name='channel']")
                noteChangerNode.SetAttributeValue("value", (map.baseChannel + remote.channelOffset).ToString()) |> ignore
            soundSlots.[remoteSlot] <- None
        
        soundSlots |> Array.filter(fun s -> s <> None) |> Array.iter(fun s -> s.Value.Remove() )

        if (map.standardRemoteAssignments) then
            // Remove unused articulation visual slots
            let slotVisuals = rootNode.SelectNodes("//member[@name='slotvisuals;]//obj[@class='USlotVisuals']") |> Seq.cast<HtmlNode> |> Seq.toArray
            for i in 0..slotVisuals.Length do
                if (soundSlots.[i] <> None) then
                    slotVisuals.[i].Remove()
        else
            rootNode.SelectNodes("//obj[@class='USlotVisuals']") |> Seq.cast<HtmlNode> |> Seq.iter(fun s -> s.Remove() )

        let output = document.DocumentNode.OuterHtml.Replace("id=", "ID=").Replace("instrumentmap", "InstrumentMap")
        File.WriteAllText(String.Format(@"Maps\{0}.expressionMap", map.name), output)

let buildExpressionMaps f = 
    let maps = parseMapFile f
    saveExpressionMaps maps

[<EntryPoint>]
let main argv = 
    buildExpressionMaps argv.[0]
    0