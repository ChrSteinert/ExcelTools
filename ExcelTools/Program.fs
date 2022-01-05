open System

open Argu

type CliArguments =
    | [<Unique; CliPrefix(CliPrefix.None)>]Depassword of ParseResults<Depassword.CliArguments>
    | [<Unique; CliPrefix(CliPrefix.None)>]Score of ParseResults<Score.CliArguments>

    interface IArgParserTemplate with
        member this.Usage =
            match this with
            | Depassword _ -> "remove password protection from any sheets of a give Excel file."
            | Score _ -> "scores given Excel files by number of cells, fomulas and so on."

[<EntryPoint>]
let main _ =
    let exiter = ProcessExiter(fun e -> if e = ErrorCode.HelpText then None else Some ConsoleColor.Red)
    let parser = ArgumentParser.Create<CliArguments> (programName = "ExcelTools", errorHandler = exiter)
    let config = parser.Parse ()

    if config.Contains Depassword
    then
        config.GetResult Depassword
        |> Depassword.depassword
    elif config.Contains Score
    then
        config.GetResult Score
        |> Score.score
    else
        config.Raise "No known command supplied!"

    0 // return an integer exit code
