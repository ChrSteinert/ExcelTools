[<RequireQualifiedAccess>]
module Score

open System.IO
open System.Xml

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet

open Argu

type CliArguments = 
    | [<Mandatory; AltCommandLine("-d")>]Directory of string
    | [<AltCommandLine("-r")>]Recursive
    | [<Mandatory; MainCommand>]ReportFile of string

    interface IArgParserTemplate with
        member this.Usage = 
            match this with
            | Directory _ -> "search directory for input files."
            | Recursive -> "search the given directory and sub directories for Excel files."
            | ReportFile _ -> "name of the output XML file."

type private Metric = 
    {
        Name : string
        Value : uint
    }

type private WorksheetReport = 
    {
        Name : string
        Metrics : Metric list
    }

type private WorkbookReport =
    {
        File : FileInfo
        WorksheetReports : WorksheetReport list
    }

let private getSheets (wb : Workbook) = 
    wb.Sheets |> Seq.cast<Sheet> |> Seq.map (fun sheet -> sheet.Name.Value, sheet.Id.Value)

let private getSheetById (wb : WorkbookPart) = 
    fun c -> 
        try 
            wb.GetPartById c
            |> fun c -> c :?> WorksheetPart
            |> Some
        with e -> 
            printfn "%O" e
            None


let private proc (file : FileInfo) =
    printfn "Reading '%s'" file.Name
    try
        use stream = file.OpenRead ()
        use package = SpreadsheetDocument.Open(stream, false)
        
        [
            for name, ws in
                package.WorkbookPart.Workbook
                |> getSheets
                |> Seq.map (fun (name, id) -> name, id |> getSheetById package.WorkbookPart) do

                    match ws with
                    | Some ws -> 
                        let definedNames = 
                            if isNull package.WorkbookPart.Workbook.DefinedNames then 0u
                            else
                                package.WorkbookPart.Workbook.DefinedNames 
                                |> Seq.cast<DefinedName> 
                                |> Seq.filter (fun c -> c.InnerText.StartsWith name)
                                |> Seq.length
                                |> uint

                        let tableParts = 
                            if ws.TableDefinitionParts |> isNull then 0u
                            else ws.TableDefinitionParts |> Seq.length |> uint


                        use wsStream = ws.GetStream(FileMode.Open, FileAccess.Read)
                        use reader = XmlReader.Create wsStream
                        let mutable cells = 0u
                        let mutable formulas = 0u
                        let mutable arrayFormulas = 0u
                        while reader.Read () do
                            if reader.NodeType = XmlNodeType.Element then
                                if reader.LocalName = "c" then cells <- cells + 1u
                                elif reader.LocalName = "f" then 
                                    if reader.MoveToAttribute "t" && reader.ReadContentAsString () = "array" then
                                        arrayFormulas <- arrayFormulas + 1u
                                    else formulas <- formulas + 1u
                                
                        Some { 
                            Name = name
                            Metrics = 
                                [ 
                                    { Name = "Cells"; Value = cells }
                                    { Name = "Formulas"; Value = formulas } 
                                    { Name = "ArrayFormulas"; Value = arrayFormulas } 
                                    { Name = "NamedRange"; Value = definedNames + tableParts } 
                                ] 
                        }
                    | None -> None
        ]
        |> List.choose id
        |> fun wsR -> Some { File = file; WorksheetReports = wsR }
    with e -> 
        printfn "%O" e
        None

let private writeMetrics (stream : Stream) (reports : WorkbookReport seq) =
    let doc = XmlDocument ()
    let __ = doc.CreateXmlDeclaration("1.0", "utf-8", "") |> doc.AppendChild
    let root = doc.CreateElement "Report" |> doc.AppendChild

    reports
    |> Seq.iter (fun wb ->
        let eWb = doc.CreateElement "File" 
        root.AppendChild eWb |> ignore
        eWb.SetAttribute("FullName", wb.File.FullName)
        eWb.SetAttribute("Name", wb.File.Name)
        eWb.SetAttribute("Length", wb.File.Length.ToString ())
        
        wb.WorksheetReports
        |> Seq.iter (fun ws ->
            let eWs = doc.CreateElement "Sheet"
            eWb.AppendChild eWs |> ignore
            eWs.SetAttribute("Name", ws.Name)

            ws.Metrics 
            |> Seq.iter (fun metric -> 
                let eM = doc.CreateElement metric.Name
                eWs.AppendChild eM |> ignore
                eM.InnerText <- sprintf "%i" metric.Value
            )
        )

    )

    stream |> doc.Save


let score (config : ParseResults<CliArguments>) =

    use outFile = new FileStream(config.GetResult ReportFile, FileMode.Create, FileAccess.Write)

    let dir, enumOption = 
        config.GetResult Directory,
        if config.Contains Recursive then 
            let r = EnumerationOptions()
            r.RecurseSubdirectories <- true
            r
        else EnumerationOptions ()
    Directory.EnumerateFiles(dir, "*.xlsx", enumOption)
    |> Seq.choose (FileInfo >> proc)
    |> writeMetrics outFile
