Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' RegFind. Utility to search text files using
''' regular expressions.
''' </summary>
''' <remarks>Written by Markus The, MOBZystems. Use any way you like.</remarks>
Module RegFind
  ' The regular expression to use
  Private _Regex As Regex

  ' Options
  Private _OptionVerbose As Boolean = False
  Private _OptionFileMode As Boolean = False
  Private _OptionNoFileNames As Boolean = False
  Private _OptionBare As Boolean = False
  Private _optionStandardInput As Boolean

  ' The number of matches across all files
  Private _MatchCount As Integer = 0
  ' The number of files that had one or more matches
  Private _FileMatchCount As Integer = 0

  ' The replace expression, if any
  Private _ReplaceExpression As String = Nothing

  ''' <summary>
  ''' Main entry point for TextEncode.
  ''' </summary>
  Public Function Main() As Integer
    ' Options
    Dim OptionRecursive As Boolean = False
    Dim OptionCaseSensitive As Boolean = False

    ' Get command line arguments
    Dim Args() As String = ParseCommandLine(Environment.CommandLine) ' Environment.GetCommandLineArgs()

    ' Parse arguments
    Dim ArgsFound As Integer = -1
    Dim SearchExpression As String = Nothing
    Dim FileSpecs As New List(Of String)

    For Each Arg As String In Args
      ' We get the name of the executable at position 0
      If ArgsFound = -1 Then
        ArgsFound = 0
      ElseIf Arg.StartsWith("-") OrElse Arg.StartsWith("/") Then
        ' Parse options
        If IsOption(Arg, "subdirs") Then
          OptionRecursive = True
        ElseIf IsOption(Arg, "verbose") Then
          _OptionVerbose = True
        ElseIf IsOption(Arg, "filemode") Then
          _OptionFileMode = True
        ElseIf IsOption(Arg, "case") Then
          OptionCaseSensitive = True
        ElseIf IsOption(Arg, "nofiles") Then
          _OptionNoFileNames = True
        ElseIf IsOption(Arg, "bare") Then
          _OptionBare = True
        ElseIf IsOption(Arg, "interactive") Then
          _optionStandardInput = True
        ElseIf IsOption(Arg, "replace") Then
          _ReplaceExpression = GetOptionArgument(Arg)
          If _ReplaceExpression Is Nothing Or _ReplaceExpression.Trim().Length = 0 Then
            Usage("Missing argument in replace option: " + Arg)
            Return 1
          End If
        ElseIf IsOption(Arg, "?") Then
          Usage("")
          Return 1
        Else
          Usage("Invalid option: " + Arg)
          Return 1
        End If
      Else
        ' Put the first argument specified into 
        ' the search expression; all others
        ' are file specifications
        Select Case ArgsFound
          Case 0
            SearchExpression = Arg
          Case Else
            FileSpecs.Add(Arg)
        End Select

        ArgsFound += 1
      End If
    Next Arg

    ' Do we have enough arguments?
    ' We need a search expression!
    If ArgsFound < 1 Then
      Usage("No search pattern specified")
      Return 1
    End If

    If _optionStandardInput Then
      If FileSpecs.Count > 0 Then
        Usage("No file masks allowed when reading from standard input")
        Return 1
      End If

      FileSpecs.Add("")
    Else
      ' If no filespecs specified, assume *
      If FileSpecs.Count = 0 Then
        FileSpecs.Add("*")
      End If
    End If

    ' Set options for the regular expression
    Dim Options As RegexOptions = RegexOptions.CultureInvariant Or RegexOptions.Compiled

    If _OptionFileMode Then
      Options = Options Or RegexOptions.Multiline Or RegexOptions.Singleline
    End If

    If Not OptionCaseSensitive Then
      Options = Options Or RegexOptions.IgnoreCase
    End If

    ' Create and validate the regular expression
    Try
      _Regex = New Regex(SearchExpression, Options Or RegexOptions.IgnoreCase)
    Catch ex As Exception
      Console.Error.WriteLine("Error in search expression:" + Environment.NewLine + ex.Message)
      Return 1
    End Try

    ' Now process all specified file masks
    For Each FileSpec As String In FileSpecs
      Try
        Dim FilePath As String
        Dim FileMask As String

        ' Special case: the file mask is empty,
        ' meaning we read from standard input
        If FileSpec = "" Then
          FilePath = ""
          FileMask = ""
        Else
          ' See if we have a wildcard file specification.
          FileMask = Path.GetFileName(FileSpec)
          If FileMask.Length > 0 AndAlso FileMask.IndexOfAny(New Char() {"*"c, "?"c}) >= 0 Then
            ' We ave a wildcard file specification.
            ' Use the directory part:
            FilePath = Path.GetDirectoryName(FileSpec)
            ' Do we have a directory part?
            If FilePath.Length > 0 Then
              ' Make it a full directory path
              FilePath = Path.GetFullPath(FilePath)
            Else
              ' Use the current directory
              FilePath = Environment.CurrentDirectory
            End If
          Else
            ' Get the full path name of the argument.
            ' This will 'root' the path if necessary and 
            ' resolve any . and .. references.
            ' The path need not physically exist!
            Dim FullPath As String = Path.GetFullPath(FileSpec)

            ' Did we specify a real directory?
            ' (If the path is not a valid directory, e.. C:\*.*,
            ' the result is simply False)
            If Directory.Exists(FullPath) Then
              FilePath = FullPath
              FileMask = "*"
            Else
              ' It is not a real directory. Therefore
              ' it must have a file name part.
              FileMask = Path.GetFileName(FullPath)

              ' No file name part? Then use "*".
              ' The path is most likely invalid!
              If IsNothing(FileMask) OrElse FileMask.Length = 0 Then
                FileMask = "*"
              End If

              ' The path before the file name is the path:
              FilePath = Path.GetDirectoryName(FullPath)
            End If
          End If

          ' Make sure the path ends in a backslash
          If Not FilePath.EndsWith(Path.DirectorySeparatorChar) Then
            FilePath += Path.DirectorySeparatorChar
          End If

        End If

        ' Now call RegFind with the path and mask.
        ' This will call itself if OptionRecursive = True
        RegFind(FilePath, FilePath, FileMask, OptionRecursive)

      Catch ex As Exception
        Console.Error.WriteLine("RegFind: " + ex.Message)
      End Try
    Next FileSpec

    If _OptionVerbose Then
      If _MatchCount = 0 Then
        Console.WriteLine("No matches.")
      Else
        Console.WriteLine(String.Format("{0} Match(es) in {1} file(s)", _MatchCount, _FileMatchCount))
      End If
    End If
    ' No error
    Return 0
  End Function

  ''' <summary>
  ''' Is the specified argument a switch of the seocified type?
  ''' </summary>
  ''' <param name="Arg"></param>
  ''' <param name="Switch">e.g. "Verbose"</param>
  ''' <returns>True if argument is of the form -V[erbose] or /V[erbose]</returns>
  Private Function IsOption(ByVal Arg As String, ByVal Switch As String) As Boolean
    If Arg.StartsWith("-") OrElse Arg.StartsWith("/") Then
      Dim Length As Integer = Math.Min(Arg.Length - 1, Switch.Length)
      If Arg.IndexOf(":"c) >= 0 Then
        Length = Math.Min(Arg.IndexOf(":"c), Length) - 1
      End If
      Return Arg.StartsWith(Arg.Substring(0, 1) + Switch.Substring(0, Length), StringComparison.InvariantCultureIgnoreCase)
    Else
      Return False
    End If
  End Function

  ''' <summary>
  ''' Get the argument of an option, i.e. the text specified after ':'
  ''' as in /r:test
  ''' </summary>
  ''' <param name="Argument"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetOptionArgument(ByVal Argument As String) As String
    Dim I As Integer = Argument.IndexOf(":"c)
    If I < 0 Then
      Return Nothing
    Else
      Return Argument.Substring(I + 1)
    End If
  End Function

  ''' <summary>
  ''' Display the usage of this program, optionally prefixed
  ''' by an error message.
  ''' </summary>
  ''' <param name="Message"></param>
  ''' <remarks></remarks>
  Private Sub Usage(ByVal Message As String)
    ' Display the error message, if any
    If Message.Length > 0 Then
      Console.Error.WriteLine(Message + Environment.NewLine)
    End If

    ' Then display usage information
    Dim V As Version = My.Application.Info.Version

    Console.WriteLine( _
      "RegFind " + V.Major.ToString() + "." + V.Minor.ToString() + "." + V.Build.ToString() + " (" + (IntPtr.Size * 8).ToString() + ") - http://www.mobzystems.com/Tools/RegFind.aspx " + Environment.NewLine + _
      Environment.NewLine + _
      "Usage:" + Environment.NewLine + _
      Environment.NewLine + _
      "RegFind [options] ""pattern"" [file-mask [file-mask] ...]" + Environment.NewLine + _
      Environment.NewLine + _
      "Searches the specified files for a regular expression pattern." + Environment.NewLine + _
      "If no file mask (e.g. *.vb or subdir\*.txt) is specified, all" + Environment.NewLine + _
      "files in the current directory are searched." + Environment.NewLine + _
      Environment.NewLine + _
      "The search pattern may be delimited by double or single quotes," + Environment.NewLine + _
      "or square brackets, e.g. [Dim S As String = """"]." + Environment.NewLine + _
      Environment.NewLine + _
      "Valid options are:" + Environment.NewLine + _
      "/r[eplace]:<replace-pattern> - replace matched text, e.g. /r:$1" + Environment.NewLine + _
      "/s[ubdirs] - search in subdirectories" + Environment.NewLine + _
      "/f[ilemode] - file mode (default is line mode)" + Environment.NewLine + _
      "/c[ase] - case sensitive (default is case insensitive)" + Environment.NewLine + _
      "/n[ofiles] - do not display file names" + Environment.NewLine + _
      "/b[are] - all output on single line (implies /n)" + Environment.NewLine + _
      "/v[erbose] - display number of matches" + Environment.NewLine + _
      "/i[nteractive] - read from standard input" _
    )
  End Sub

  ''' <summary>
  ''' Perform the actual search
  ''' </summary>
  ''' <param name="RootDir">The root directory of this operation</param>
  ''' <param name="FilePath">The path to process</param>
  ''' <param name="FileMask">The file mask (e.g. "*")</param>
  ''' <param name="Recursive">Process subdirectories</param>
  Private Sub RegFind( _
    ByVal RootDir As String, _
    ByVal FilePath As String, _
    ByVal FileMask As String, _
    ByVal Recursive As Boolean _
)
    ' Did we specify stdin as input file?
    If FilePath = "" Then
      ' Handle stdin in file or line mode
      If _OptionFileMode Then
        HandleFileMode(Console.In.ReadToEnd(), Nothing, Nothing)
      Else
        HandleLineMode(ReadLines(), Nothing, Nothing)
      End If
    Else
      Dim DirInfo As New DirectoryInfo(FilePath)
      Dim Files() As FileInfo = DirInfo.GetFiles(FileMask)

      For Each F As FileInfo In Files
        ' Working in file mode?
        If _OptionFileMode Then
          ' Handle the file in file mode
          HandleFileMode(File.ReadAllText(F.FullName, Encoding.Default), RootDir, F)
        Else
          ' Match on every line
          HandleLineMode(File.ReadAllLines(F.FullName, Encoding.Default), RootDir, F)
        End If
      Next F

      ' Process subdirectories if required
      If Recursive Then
        Dim Dirs() As DirectoryInfo = DirInfo.GetDirectories()
        For Each Dir As DirectoryInfo In Dirs
          RegFind(RootDir, Dir.FullName, FileMask, Recursive)
        Next Dir
      End If
    End If
  End Sub

  ''' <summary>
  ''' Read the contents of StdIn to a string array
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ReadLines() As String()
    Dim Lines As New List(Of String)

    ' Read a line from standard input
    ' until we read Nothing
    Do
      Dim L As String = Console.In.ReadLine()
      If L IsNot Nothing Then
        Lines.Add(L)
      Else
        Exit Do
      End If
    Loop

    Return Lines.ToArray()
  End Function

  ''' <summary>
  ''' Perform matching on a string
  ''' </summary>
  ''' <param name="Contents"></param>
  ''' <param name="RootDir"></param>
  ''' <param name="F"></param>
  ''' <remarks></remarks>
  Private Sub HandleFileMode(ByVal Contents As String, ByVal RootDir As String, ByVal F As FileInfo)
    Dim MC As MatchCollection = _Regex.Matches(Contents)

    If MC.Count > 0 Then
      _FileMatchCount += 1
      For Each M As Match In MC
        _MatchCount += 1
        If Not _OptionBare AndAlso Not _OptionNoFileNames AndAlso Not F Is Nothing Then
          Console.WriteLine(F.FullName.Substring(RootDir.Length) + ":")
        End If
        Dim Output As String
        If _ReplaceExpression IsNot Nothing Then
          Output = _Regex.Replace(M.Groups(0).Value, _ReplaceExpression)
        Else
          Output = M.Groups(0).Value
        End If
        If _OptionBare Then
          Console.Write(Output)
        Else
          Console.WriteLine(Output)
        End If
      Next M
    End If
  End Sub

  ''' <summary>
  ''' Perform matching on a list of strings
  ''' </summary>
  ''' <param name="Lines"></param>
  ''' <param name="RootDir"></param>
  ''' <param name="F"></param>
  ''' <remarks></remarks>
  Private Sub HandleLineMode(ByVal Lines() As String, ByVal RootDir As String, ByVal F As FileInfo)
    Dim LineNumber As Integer = 0
    For Each Line As String In Lines
      LineNumber += 1
      Dim MC As MatchCollection = _Regex.Matches(Line)
      If MC.Count > 0 Then
        _FileMatchCount += 1
        For Each M As Match In MC
          _MatchCount += 1
          If Not _OptionBare AndAlso Not _OptionNoFileNames AndAlso Not F Is Nothing Then
            Console.Write(F.FullName.Substring(RootDir.Length) + " (" + LineNumber.ToString() + "): ")
          End If
          Dim Output As String = Line
          If _ReplaceExpression IsNot Nothing Then
            Output = _Regex.Replace(Output, _ReplaceExpression)
          End If
          If _OptionBare Then
            Console.Write(Output)
          Else
            Console.WriteLine(Output)
          End If
          ' TODO: skip rest of matches when in full line mode
        Next M
      End If
    Next Line
  End Sub

  ''' <summary>
  ''' Parse the command line, allowing for double quotes, single quotes
  ''' and square brackets as command delimiters
  ''' </summary>
  ''' <param name="CommandLine"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ParseCommandLine(ByVal CommandLine As String) As String()
    Dim Arguments As List(Of String) = New List(Of String)()

    ' Loop over the characters in the command line
    Dim I As Integer = 0
    Do
      Arguments.Add(EatArgument(CommandLine, I))

      ' Eat whitespace
      Do While I < CommandLine.Length AndAlso CommandLine.Chars(I) = " "c
        I += 1
      Loop
    Loop While I < CommandLine.Length

    Return Arguments.ToArray()
  End Function

  ''' <summary>
  ''' Eat an argument from the specified command line at
  ''' the specified position. Allow single quote, double quote
  ''' and saure brackets as delimiters. If no delimiter is found,
  ''' the argument ends at the next space or end of line.
  ''' </summary>
  ''' <param name="CommandLine"></param>
  ''' <param name="Position"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function EatArgument(ByVal CommandLine As String, ByRef Position As Integer) As String
    ' Is the first character a double quote?
    If CommandLine.Chars(Position) = """"c Or CommandLine.Chars(Position) = "'"c Or CommandLine.Chars(Position) = "["c Then
      Dim EndChar As Char
      If CommandLine.Chars(Position) = "["c Then
        EndChar = "]"c
      Else
        EndChar = CommandLine.Chars(Position)
      End If

      ' Eat all characters up to the next double quote
      Dim LastPosition As Integer = Array.IndexOf(CommandLine.ToCharArray(), EndChar, Position + 1)
      If LastPosition < 0 Then
        Throw New ApplicationException("Unterminated argument starting with " + CommandLine.Substring(Position))
      End If

      Dim Arg As String = CommandLine.Substring(Position + 1, LastPosition - Position - 1)
      Position = LastPosition + 1
      Return Arg
    Else
      ' Find a space to terminate the argument
      Dim LastPosition As Integer = Array.IndexOf(CommandLine.ToCharArray(), " "c, Position + 1)
      If LastPosition < 0 Then
        ' No more spaces - use rest of string
        LastPosition = CommandLine.Length
      End If

      Dim Arg As String = CommandLine.Substring(Position, LastPosition - Position)
      Position = LastPosition + 1
      Return Arg
    End If
  End Function
End Module
