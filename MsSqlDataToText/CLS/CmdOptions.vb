Public Class CmdOptions

    <CommandLine.Option('cfg', "config", Required := true,
    HelpText:="Configuration XML.")>
    Public Property InputFiles As IEnumerable(Of String)

End Class
