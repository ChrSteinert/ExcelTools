open System
open System.IO
open System.IO.Compression
open System.Xml

open Argu

type CliArguments = 
    // | [<AltCommandLine("-i")>] Inplace
    | [<MainCommand; ExactlyOnce; Last>] Files of string list

    interface IArgParserTemplate with
        member this.Usage = 
            match this with
            // | Inplace -> "change the files directly, instead of making copies."
            | Files _ -> "the files."


let findWorksheets (contentTypesXml : XmlReader) = 
    let reader = contentTypesXml
    [
        while reader.Read () do
            if reader.LocalName = "Override" && reader.GetAttribute "ContentType" = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" then
                yield reader.GetAttribute "PartName"
    ]     

let removeSheetProtection (sheetDoc : XmlDocument) =    
    let nodes = sheetDoc.GetElementsByTagName "sheetProtection"
    if nodes.Count = 1 then
        
        (nodes.ItemOf 0).ParentNode.RemoveChild(nodes.ItemOf 0)
        |> ignore
        printfn "Removed sheet protection."
    else
        printfn "Sheet was not protected."


[<EntryPoint>]
let main _ =
    let exiter = ProcessExiter(fun _ -> Some ConsoleColor.Red)
    let parser = ArgumentParser.Create<CliArguments> (programName = "ExcelTools", errorHandler = exiter)
    let config = parser.Parse ()
    config.GetResult Files
    |> List.iter (fun file ->
        let file = file |> FileInfo
        if file.Exists then
            printf "Found file '%s'" file.FullName
            try
                use zip = ZipFile.OpenRead(file.FullName)
                let fileName = sprintf "%sunprotected%s" file.FullName.[0..file.FullName.Length - file.Extension.Length] file.Extension
                try
                    use outFile =     
                        printfn " … will write to '%s'" fileName
                        ZipFile.Open(fileName, ZipArchiveMode.Create)
                    
                    let isSheet = 
                        let sheets = 
                            let entry = zip.GetEntry("[Content_Types].xml")
                            use stream = entry.Open ()
                            use reader = XmlReader.Create stream
                            reader |> findWorksheets |> List.map (fun c -> c.[1..]) 
                        fun partName -> List.contains partName sheets
                    
                    zip.Entries
                    |> Seq.iter (fun entry -> 
                        if entry.FullName |> isSheet then
                            printfn "Processing %s" entry.FullName
                            use inStream = entry.Open ()
                            let doc = 
                                let doc = XmlDocument()
                                doc.Load inStream
                                doc
                            doc |> removeSheetProtection

                            let newEntry = outFile.CreateEntry entry.FullName
                            use outStream = newEntry.Open ()
                            doc.Save(outStream)
                        else
                            printfn "Copying %s" entry.FullName
                            use inStream = entry.Open ()
                            let newEntry = outFile.CreateEntry entry.FullName
                            use outStream = newEntry.Open ()
                            inStream.CopyTo outStream
                    )

                    printfn "File '%s' processed\n" file.FullName
                with e -> printfn "Could not write target file '%s' with error: %s\n" fileName e.Message
            with e -> printfn "\nCould not process file '%s'.\n\tIs it an Excel file?\n\tIs it a password protected file?\n" file.FullName
        else printfn "File %O could not be found!\n" file
    )

    0 // return an integer exit code
