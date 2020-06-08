'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Convert
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal

'This module contains the user interface routines.
Public Module InterfaceModule
   'The Microsoft Windows API constants and functions used by this program.
   <DllImport("Kernel32.dll", SetLastError:=True)> Private Function GetBinaryTypeA(ByVal lpApplicationName As String, ByRef lpBinaryType As Integer) As Integer
   End Function

   Private Const ERROR_SUCCESS As Integer = &H0%
   Private Const SCS_32BIT_BINARY As Integer = &H0%
   Private Const SCS_DOS_BINARY As Integer = &H1%
   Private Const SCS_WOW_BINARY As Integer = &H2%
   Private Const SCS_PIF_BINARY As Integer = &H3%
   Private Const SCS_POSIX_BINARY As Integer = &H4%
   Private Const SCS_OS216_BINARY As Integer = &H5%
   Private Const SCS_64BIT_BINARY As Integer = &H6%

   'This enumeration lists the command line arguments supported by this program.
   Private Enum ArgumentsE As Integer
      Path       'The path argument.
      Range      'The range argument.
      Options    'The options argument.
      Pause      'The pause argument.
   End Enum

   'This structure defines this program's command line arguments.
   Private Structure CommandLineArgumentsStr
      Public Count As String           'Defines the number of bytes to be disassembled.
      Public InputFile As String       'Defines the file containing the code to be disassembled.
      Public Options As String         'Defines the options used to control the disassembly process.
      Public PauseInterval As Integer  'Defines the number of lines to be displayed before pausing.
      Public StartPosition As String   'Defines the position at which to start disassembling.
   End Structure

   Private Const COUNT_DELIMITER As Char = "-"c  'Defines the delimiter for the number of bytes to be disassembled and start position.
   Private Const OPTION_HEADER As Char = "h"c    'Defines the "interpret as MS-DOS executable header" option specifier.
   Private Const OPTION_TYPE As Char = "t"c      'Defines the "display binary type" option specifier.
   Private Const OPTION_WAIT As Char = "w"c      'Defines the "wait before quitting" option specifier.

   Private WithEvents Disassembler As New DisassemblerClass   'Contains the reference to the disassembler.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim Arguments As CommandLineArgumentsStr = ProcessCommandLine()
         Dim Code As New List(Of Byte)
         Dim Position As Integer = TextToNumber(Arguments.StartPosition)

         ExitCode = ERROR_SUCCESS

         If Arguments.InputFile = Nothing Then
            DisplayHelpAndInformation()
         Else
            Code = New List(Of Byte)(File.ReadAllBytes(Arguments.InputFile))

            If Code IsNot Nothing Then
               If Arguments.Options.Contains(OPTION_HEADER) Then
                  DisplayMsDosExeHeader(Code, Position, PauseInterval:=Arguments.PauseInterval)
               ElseIf Arguments.Options.Contains(OPTION_TYPE) Then
                  WriteToConsole(GetFileType(Arguments.InputFile))
               Else
                  For Each Line As String In DisasembledCode(Code, Position, Arguments)
                     If WriteToConsole(Line, , PauseInterval:=Arguments.PauseInterval) = ToChar(ConsoleKey.Escape) Then Exit For
                  Next Line
               End If
            End If
         End If

         If Arguments.Options.Contains(OPTION_WAIT) Then Console.ReadKey(intercept:=True)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure checks whether an error has occurred during the most recent Windows API call.
   Private Function CheckForError(Optional ReturnValue As Object = Nothing, Optional ExtraInformation As String = "") As Object
      Try
         Dim Description As String = Nothing
         Dim ErrorCode As Integer = GetLastWin32Error()
         Dim Message As String = Nothing

         If Not ErrorCode = ERROR_SUCCESS Then
            Description = New Win32Exception(ErrorCode).Message
            If Description.Contains("%1") Then Description = Description.Replace("%1", ExtraInformation)

            Message = ($"API error code: {ErrorCode} - {Description}{NewLine}Return value: {ReturnValue}{NewLine}")
            WriteToConsole(Message)
         End If

         Return ReturnValue
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return ReturnValue
   End Function

   'This procedure disassembles the specified code and returns the resulting lines of code.
   Private Function DisasembledCode(Code As List(Of Byte), Position As Integer, Arguments As CommandLineArgumentsStr) As List(Of String)
      Try
         Dim Count As Integer = TextToNumber(Arguments.Count)
         Dim Disassembly As New List(Of String)
         Dim EndPosition As Integer = If(Count = &H0%, Code.Count, Position + Count)
         Dim HexadecimalCode As String = Nothing
         Dim Instruction As String = Nothing
         Dim PreviousPosition As New Integer

         With Disassembler
            Do Until Position >= EndPosition
               PreviousPosition = Position + &H1%
               Instruction = .Disassemble(Code, Position)
               HexadecimalCode = .BytesToHexadecimal(.GetBytes(Code, PreviousPosition - &H1%, (Position - PreviousPosition) + &H1%), NoPrefix:=True, Reverse:=False)
               Disassembly.Add($"{(PreviousPosition - &H1%):X8} {HexadecimalCode.PadRight(25, " "c)}{Instruction}")
            Loop
         End With

         Return Disassembly
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure displays the help and information for this program.
   Private Sub DisplayHelpAndInformation()
      Try
         With My.Application.Info
            WriteToConsole($"{ .Title} - v{ .Version} - by: { .CompanyName} { .Copyright}")
            WriteToConsole()
            WriteToConsole("Dasm8086 [PATH]*[OFFSET[-COUNT]]*[OPTIONS]*[PAUSE]")
            WriteToConsole("PATH - The file to be disassembled.")
            WriteToConsole("OFFSET - The position at which to start disassembling in (hexa)decimals.")
            WriteToConsole("COUNT - The number of bytes to be disassambled in (hexa)decimals. 0 = all bytes.")
            WriteToConsole("OPTIONS - A sequence of characters that enable the following options:")
            WriteToConsole("          h = Interprete and display the data as an MS-DOS executable header.")
            WriteToConsole("          t = Display the binary type for the specified file.")
            WriteToConsole("          w = Wait for the user to press a key before quiting.")
            WriteToConsole("PAUSE - The number lines to display before pausing specified in decimals.")
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the contents of the specified MS-DOS executable header.
   Private Sub DisplayMsDosExeHeader(Code As List(Of Byte), ByRef Position As Integer, PauseInterval As Integer)
      Try
         Dim RelocationCount As String = Nothing

         With Disassembler
            WriteToConsole($"Signature               00  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Last page size          02  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"File size **            04  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            RelocationCount = .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)
            WriteToConsole($"Relocation count        06  {RelocationCount}")
            WriteToConsole($"Header size *           08  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Required memory *       0A  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Prefered memory *       0C  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Initial SS              0E  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Initial SP              10  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"PGM checksum            12  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Initial IP              14  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Unrelocated CS          16  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Relocation table offset 18  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole($"Overlay number          1A  { .BytesToHexadecimal(.GetBytes(Code, Position, Length:=2), NoPrefix:=True)}")
            WriteToConsole()
            WriteToConsole($"*= Size specified in paragraphs of { .HEXADECIMAL_PREFIX}10 bytes.")
            WriteToConsole($"**= Size specified in pages of { .HEXADECIMAL_PREFIX}200 bytes.")
            WriteToConsole()

            WriteToConsole("Relocation items:")
            For Item As Integer = &H1% To ToInt32(RelocationCount, fromBase:=16)
               If WriteToConsole(.BytesToHexadecimal(.GetBytes(Code, Position, Length:=4)), , PauseInterval) = ToChar(ConsoleKey.Escape) Then Exit For
            Next Item
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure returns the binary type description for the specified file.
   Private Function GetFileType(Path As String) As String
      Try
         Dim BinaryType As New Integer

         If CBool(CheckForError(GetBinaryTypeA(Path, BinaryType), Path)) Then
            Select Case BinaryType
               Case SCS_32BIT_BINARY
                  Return "32-bit Windows"
               Case SCS_DOS_BINARY
                  Return "MS-DOS"
               Case SCS_WOW_BINARY
                  Return "16-bit Windows"
               Case SCS_PIF_BINARY
                  Return "PIF for MS-DOS applications"
               Case SCS_POSIX_BINARY
                  Return "Posix"
               Case SCS_OS216_BINARY
                  Return "16-bit OS/2"
               Case SCS_64BIT_BINARY
                  Return "64-bit Windows"
            End Select
         End If

         Return "unknown"
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure handles any errors that occur.
   Private Sub HandleError(ExceptionO As Exception) Handles Disassembler.HandleError
      Try
         Dim ErrorCode As Integer = Microsoft.VisualBasic.Err.Number

         With Console.Error
            .WriteLine()
            .Write("Error: ")
            If Not ErrorCode = ERROR_SUCCESS Then .Write($"{ErrorCode} - ")
            .WriteLine(ExceptionO.Message)
            .WriteLine()
         End With

         ExitCode = ErrorCode
      Catch
         [Exit](CInt(False))
      End Try
   End Sub

   'This procedure returns the processed command line arguments.
   Private Function ProcessCommandLine() As CommandLineArgumentsStr
      Try
         Dim Arguments As New CommandLineArgumentsStr With {.Count = "0x00", .InputFile = Nothing, .Options = "", .PauseInterval = 0, .StartPosition = "0x00"}
         Dim ArgumentsList As New List(Of String)

         If My.Application.CommandLineArgs.Any Then
            ArgumentsList.AddRange(My.Application.CommandLineArgs.First.Split("*"c))

            Do Until ArgumentsList.Count > ArgumentsE.Pause
               ArgumentsList.Add("")
            Loop

            With Arguments
               .InputFile = ArgumentsList(ArgumentsE.Path).Trim
               If .InputFile.StartsWith("""") Then .InputFile = .InputFile.Substring(1)
               If .InputFile.EndsWith("""") Then .InputFile = .InputFile.Substring(0, .InputFile.Length - 1)
               If ArgumentsList(ArgumentsE.Range).Contains(COUNT_DELIMITER) Then
                  .StartPosition = ArgumentsList(ArgumentsE.Range).Substring(0, ArgumentsList(ArgumentsE.Range).IndexOf(COUNT_DELIMITER)).Trim
                  .Count = ArgumentsList(ArgumentsE.Range).Substring(ArgumentsList(ArgumentsE.Range).IndexOf(COUNT_DELIMITER) + 1).Trim
               Else
                  .StartPosition = ArgumentsList(ArgumentsE.Range).Trim
               End If
               .Options = ArgumentsList(ArgumentsE.Options).Trim
               Integer.TryParse(ArgumentsList(ArgumentsE.Pause), .PauseInterval)
               If .StartPosition.ToLower.StartsWith(Disassembler.HEXADECIMAL_PREFIX.ToLower) Then .StartPosition = ToInt32(.StartPosition.Substring(Disassembler.HEXADECIMAL_PREFIX.Length), fromBase:=16).ToString()
            End With
         End If

         Return Arguments
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified text to a number.
   Private Function TextToNumber(Text As String) As Integer
      Try
         If Text Is Nothing OrElse Text = Nothing Then
            Return 0
         ElseIf Text.ToLower.StartsWith(Disassembler.HEXADECIMAL_PREFIX) Then
            Return ToInt32(Text.Substring(Disassembler.HEXADECIMAL_PREFIX.Length), fromBase:=16)
         Else
            Return ToInt32(Text)
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure writes the specified text to the console.
   Private Function WriteToConsole(Optional Text As String = Nothing, Optional NoNewLine As Boolean = False, Optional PauseInterval As Integer = 0) As String
      Try
         Static LinesDisplayed As Integer = 0

         Console.Write(Text)
         If Not NoNewLine Then
            Console.WriteLine()
            LinesDisplayed += 1
         End If

         Return If(PauseInterval > 0 AndAlso LinesDisplayed Mod PauseInterval = 0, Console.ReadKey(intercept:=True).KeyChar, Nothing)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
