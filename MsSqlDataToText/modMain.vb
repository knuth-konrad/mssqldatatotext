' Idea shamelessly stolen from https://stackoverflow.com/questions/30791482/sql-server-management-studio-2012-export-all-tables-of-database-as-csv
' Imports CommandLine

Imports System.Data
Imports System.Data.SqlClient

Imports libBAUtil.ConsoleUtil
Imports libBAUtil.FilesystemUtil
Imports libBAUtil.StringUtil
Imports libBAUtil.TextFileUtil
Imports libBAUtil.DBTools.SQL.SqlDBUtil
Imports libXmlConfig.Tools.Xml.Configuration

Module modMain

   Sub Main(ByVal cmdLineArgs() As String)

      'CommandLine.Parser.Default.ParseArguments(Of CmdOptions)(cmdLineArgs) _
      '        .WithParsed(Function(opts As CmdOptions) RunOptionsAndReturnExitCode(opts)) _
      '        .WithNotParsed(Function(errs As IEnumerable(Of [Error])) 1)

      cBAUtilCC.ConHeadline("MsSqlDataToText")
      cBAUtilCC.ConCopyright()
      BlankLine()

      Console.WriteLine("Using configuration file: {0}", cmdLineArgs(0))

      ' Parse the configuration file
      Dim xmlConfig As New XmlConfig(cmdLineArgs(0))

      If xmlConfig.Init() <> XmlConfig.eCfgXMLResult.xmlSuccess Then

         Console.WriteLine("Can't parse XML config")
         Console.ReadLine()

      Else

         Console.WriteLine("Total records exported: {0}", ExportAllData(xmlConfig).ToString)

      End If

   End Sub

   Function ExportAllData(ByVal xmlConfig As XmlConfig) As Int32

      Dim sConnectionstring As String = ""
      sConnectionstring = xmlConfig.DefaultConfig.Sections.GetSection("Database").Entrys.GetEntry("ConnectionString").Value.ToString

      Dim sTableQuery As String = ""
      sTableQuery = xmlConfig.DefaultConfig.Sections.GetSection("Database").Entrys.GetEntry("TableQuery").Value.ToString

      ' Retrieve all tables from the database schema
      Dim dbCon As New SqlConnection(sConnectionstring)
      Dim dbCmd As New SqlCommand(sTableQuery, dbCon)

      ' Load up the tables and schemas in a dataset
      Dim dbAdapter As New SqlDataAdapter(dbCmd)
      Dim dbDataset As New DataSet
      dbAdapter.Fill(dbDataset)

      ' Loop through all tables and export a CSV of the Table Data
      Dim lRet, lTotal As Int32
      For Each row As DataRow In dbDataset.Tables(0).Rows

         lRet = ExportTableData(xmlConfig.DefaultConfig, dbCon, row)
         lTotal += lRet

         Console.WriteLine(" - Records exported: " & lRet.ToString)
         ' Console.WriteLine(ExportTableData(xmlConfig.DefaultConfig, row))
         ' Console.WriteLine("dbDataset.Tables(0).Rows: " & row(0).ToString & "_" & row(1).ToString)

         BlankLine()

      Next

      If dbCon.State = ConnectionState.Closed Then
         dbCon.Open()
      End If

      Return lTotal

   End Function

   Function ExportTableData(ByVal xmlCfg As IXmlCfg, ByVal dbCon As SqlConnection, ByVal row As DataRow) As Int32
      ' Export the data of a specific table

      ' Create the text file name, based upon the table's name
      Dim sOutFile As String = ""
      sOutFile = CreateFilename(xmlCfg, row)
      Console.WriteLine("-- Outfile: " & sOutFile)

      Dim sTablename As String = row(0).ToString & "." & row(1).ToString
      Dim lRows As Int32

      If DoSkipTable(xmlCfg, sTablename) = True Then

         Console.WriteLine(" - Skipping table: " & sTablename)
         Return 0

      Else

         Dim sQuery As String
         sQuery = System.String.Format("SELECT COUNT(0) AS RecCount FROM [{0}].[{1}]", row(0), row(1))

         If dbCon.State = ConnectionState.Closed Then
            dbCon.Open()
         End If
         Dim dbCmd As New SqlCommand(sQuery, dbCon)

         ' *** Count the number of rows in that table
         Dim dbReader As SqlDataReader = dbCmd.ExecuteReader()
         Do While dbReader.Read
            lRows = libBAUtil.DBTools.SQL.SqlDBUtil.GetDBInteger(dbReader, "RecCount")
         Loop
         dbReader.Close()

         ' Skip empty table alltogether?
         If lRows = 0 Then
            If (xmlCfg.Sections.GetSection("Export").Entrys.HasEntry("SkipEmptyTables") = True) AndAlso
               (CBool(xmlCfg.Sections.GetSection("Export").Entrys.GetEntry("SkipEmptyTables").Value) = True) Then
               Console.WriteLine(" - Skipping empty table: " & sTablename)
               Return lRows
            End If
         End If


         ' *** Do the actual data export of the table
         ' Retrieve the column names
         'sQuery = System.String.Format("SELECT * FROM [{0}].[{1}]", row(0), row(1))
         sQuery = System.String.Format("SELECT TOP 1 * FROM [{0}].[{1}]", row(0), row(1))

         ' A value of 0 indicates no limit (an attempt to execute a command will wait indefinitely).
         dbCmd.CommandText = sQuery
         dbCmd.CommandTimeout = 0
         Dim dbAdapter As New SqlDataAdapter(dbCmd)

         Dim dbDataset As New DataSet
         dbAdapter.Fill(dbDataset)

         Dim dbDataTable As DataTable = dbDataset.Tables(0)

         ' *** Create the text file header, if configured to do so
         CreateTextFileHeader(xmlCfg, sOutFile, sTablename, dbDataTable.Columns)

         ' *** Export the actual data
         ' Column delimiter
         Dim sDelim As String = xmlCfg.Sections.GetSection("Export").Entrys.GetEntry("ColumnDelimiter").Value.ToString

         sQuery = System.String.Format("SELECT * FROM [{0}].[{1}]", row(0), row(1))
         'sQuery = System.String.Format("SELECT TOP 100 * FROM [{0}].[{1}]", row(0), row(1))
         dbCmd.CommandText = sQuery

         dbReader = dbCmd.ExecuteReader()
         Do While dbReader.Read

            Dim txtLine As String = ""

            For i As Int32 = 0 To dbReader.FieldCount - 1

               ' Skip column?
               If Not DoSkipColumn(xmlCfg, sTablename & "." & dbReader.GetName(i)) Then

                  Dim sTemp As String = ""

                  With dbReader

                     If Not .IsDBNull(i) And (.GetFieldType(i) Is GetType(Byte())) Then ' .DataType.ToString = "System.Byte[]"
                        'Dim abyt() As Byte = CType(dbRow.Item(dbCol), Byte())
                        Dim abyt() As Byte = CType(.GetValue(i), Byte())
                        ' Console.WriteLine(" dbRow.Item: {0}", System.Text.ASCIIEncoding.ASCII.GetString(CType(dbRow.Item(dbCol), Byte())))
                        For Each b As Byte In abyt
                           sTemp &= b.ToString
                        Next
                        'Console.WriteLine(" dbRow.Item: {0}", sTemp)
                     Else
                        'Console.WriteLine(" dbRow.Item: {0}", dbRow.Item(dbCol).ToString)
                        sTemp = ReadNullAsEmptyString(dbReader, i)
                     End If

                     If txtLine.Length > 0 Then
                        txtLine &= sDelim
                     End If

                     txtLine &= EnQuote(sTemp)

                  End With

               End If

            Next

            TxtWriteLine(sOutFile, txtLine)

         Loop
         dbReader.Close()


         'For Each dbRow As DataRow In dbDataTable.Rows

         '   Dim txtLine As String = ""

         '   For Each dbCol As DataColumn In dbDataTable.Columns

         '      'Console.WriteLine("dbCol.Name: {0}", dbCol.ColumnName)
         '      'Console.WriteLine("dbCol.Type: {0}", dbCol.DataType.ToString)

         '      Dim sTemp As String = ""

         '      If Not dbRow.IsNull(dbCol) And dbCol.DataType.ToString = "System.Byte[]" Then
         '         Dim abyt() As Byte = CType(dbRow.Item(dbCol), Byte())
         '         ' Console.WriteLine(" dbRow.Item: {0}", System.Text.ASCIIEncoding.ASCII.GetString(CType(dbRow.Item(dbCol), Byte())))
         '         For Each b As Byte In abyt
         '            sTemp &= b.ToString
         '         Next
         '         'Console.WriteLine(" dbRow.Item: {0}", sTemp)
         '      Else
         '         'Console.WriteLine(" dbRow.Item: {0}", dbRow.Item(dbCol).ToString)
         '         sTemp = dbRow.Item(dbCol).ToString
         '      End If

         '      If txtLine.Length > 0 Then
         '         txtLine &= sDelim
         '      End If

         '      txtLine &= EnQuote(sTemp)

         '   Next

         '   TxtWriteLine(sOutFile, txtLine)

         'Next

         'dbDataTable.WriteXml(sOutFile, XmlWriteMode.IgnoreSchema, False)

      End If

      Return lRows

   End Function

   Function CreateFilename(ByVal xmlCfg As IXmlCfg, ByVal row As DataRow, Optional ByVal sExtension As String = ".csv") As String

      ' Construct the output file name
      Dim sRet As String = ""

      sRet = NormalizePath(xmlCfg.Sections.GetSection("Export").Entrys.GetEntry("DestinationPath").Value.ToString) & row(0).ToString & "_" & row(1).ToString & sExtension
      Return sRet

   End Function

   Function CreateTextFileHeader(xmlCfg As IXmlCfg, ByVal exportFile As String, ByVal tableName As String, ByVal dbCols As DataColumnCollection) As Boolean

      ' Safe guard - write header at all?
      If xmlCfg.Sections.GetSection("Export").Entrys.HasEntry("ColumnNameAsFirstLine") Then
         Dim o As IXmlCfgEntry
         o = xmlCfg.Sections.GetSection("Export").Entrys.GetEntry("ColumnNameAsFirstLine")
         If CBool(o.Value) = False Then
            Return True
         End If
      End If

      Dim sDelim As String = ""
      Dim sHdrText As String = ""

      sDelim = xmlCfg.Sections.GetSection("Export").Entrys.GetEntry("ColumnDelimiter").Value.ToString

      For Each dbCol As DataColumn In dbCols

         If Not DoSkipColumn(xmlCfg, tableName & "." & dbCol.ColumnName) Then
            If sHdrText.Length > 0 Then
               sHdrText &= sDelim
            End If

            sHdrText &= EnQuote(dbCol.ColumnName)
         End If

      Next

      ' Write the column names as the first line.
      ' As this is the first line, recreate the file
      TxtWriteLine(exportFile, sHdrText, False)

      Return True

   End Function

   Function DoSkipTable(ByVal xmlCfg As IXmlCfg, ByVal tableName As String) As Boolean

      ' Determine if a table should be skipped
      For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection("SkipColumns").Entrys.CfgEntrys
         'Console.WriteLine("Config value: {0}, parameter: {1}", o.Value.ToString, tableName)
         If o.Value.ToString = tableName & ".*" Then
            Return True
         End If
      Next

      Return False

   End Function

   Function DoSkipColumn(ByVal xmlCfg As IXmlCfg, ByVal columnName As String) As Boolean

      ' Determine if a certain column should be skipped
      For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection("SkipColumns").Entrys.CfgEntrys
         'Console.WriteLine("Config value: {0}, parameter: {1}", o.Value.ToString, tableName)
         If o.Value.ToString = columnName Then
            Return True
         End If
      Next

      Return False

   End Function

End Module
