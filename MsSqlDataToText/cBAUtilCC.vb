Imports libBAUtil.StringUtil

''' <summary>
''' Helper methods for console applications.
''' </summary>
Public Class cBAUtilCC

   ' Defaults
   Private Const APP_AUTHOR As String = "Knuth Konrad"
   Private Const APP_COMPANY As String = "STA Travel GmbH"

   Public Structure PrgVersionSTRUCT
      Public Major As Int32
      Public Minor As Int32
      Public Revision As Int32
      Public Build As Int32
   End Structure

   Private msApplication As String = ""
   Private msAuthor As String = ""
   Private msCompany As String = ""

   Private mudtVersion As PrgVersionSTRUCT

   Public Property Application As String
      Get
         Return msApplication
      End Get
      Set(value As String)
         msApplication = value
      End Set
   End Property

   Public Property Author As String
      Get
         Return msAuthor
      End Get
      Set(value As String)
         msAuthor = value
      End Set
   End Property

   Public Property Company As String
      Get
         Return msCompany
      End Get
      Set(value As String)
         msCompany = value
      End Set
   End Property

   Public Property ProgramVersion As PrgVersionSTRUCT
      Get
         Return mudtVersion
      End Get
      Set(value As PrgVersionSTRUCT)
         mudtVersion = value
      End Set
   End Property

   Public Overloads Shared Sub ConHeadline(ByVal sAppName As String, ByVal iVersionMajor As Integer,
      ByVal iVersionMinor As Integer, Optional ByVal iVersionRevision As Int32 = 0, Optional iVersionBuild As Int32 = 0)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & sAppName & " v" &
                        iVersionMajor.ToString & "." &
                        iVersionMinor.ToString & "." &
                        iVersionRevision.ToString & "." &
                        iVersionBuild.ToString &
                        " *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   Public Overloads Shared Sub ConHeadline(ByVal sProgName As String)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & sProgName & " *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   Public Overloads Shared Sub ConHeadline(ByVal sProgName As String, ByVal iMajorVersion As Integer)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & sProgName & " v" & CType(iMajorVersion, String) & ".0 *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   Public Overloads Shared Sub ConCopyright()

      Console.WriteLine("Copyright " & Chr(169) & " " & DateTime.Now.Year.ToString & " by STA Travel GmbH. All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   Public Overloads Shared Sub ConCopyright(ByVal sYear As String, ByVal sCompanyname As String)

      Console.WriteLine("Copyright " & Chr(169) & " " & sYear & " by " & sCompanyname & ". All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   Public Overloads Shared Sub ConCopyright(ByVal sCompanyname As String)

      Console.WriteLine("Copyright " & Chr(169) & " " & DateTime.Now.Year.ToString & " by " & sCompanyname & ". All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   Public Sub New()

      MyBase.New

      With Me
         .Author = APP_AUTHOR
         .Company = APP_COMPANY
      End With

   End Sub

   Public Sub New(ByVal companyName As String, ByVal authorName As String, ByVal udtVersion As PrgVersionSTRUCT)

      MyBase.New

      With Me
         .Author = authorName
         .Company = companyName
         .ProgramVersion = udtVersion
      End With
   End Sub

End Class
