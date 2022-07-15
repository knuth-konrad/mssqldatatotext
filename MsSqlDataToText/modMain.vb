' Idea shamelessly stolen from https://stackoverflow.com/questions/30791482/sql-server-management-studio-2012-export-all-tables-of-database-as-csv
' Imports CommandLine

Imports System.Data
Imports System.Data.SqlClient

Imports libBAUtil.ConsoleHelper
Imports libBAUtil.FilesystemHelper
Imports libBAUtil.StringHelper
Imports libBAUtil.TextFileUtil
Imports libBAUtil.DBTools.SQL.SqlDBUtil
Imports libXmlConfig.Tools.Xml.Configuration


Module modMain

#Region "Declarations"
   ' Sections in the XML configuration file
   Public Const CFG_DATABASE As String = "Database"
   Public Const CFG_EXPORT As String = "Export"
   Public Const CFG_EXPORTCOLUMNS As String = "ExportColumns"
   Public Const CFG_SKIPCOLUMNS As String = "SkipColumns"
   Public Const CFG_TABLESELECT As String = "TableSelect"
   Public Const CFG_TABELNAMESFROMDB As String = "TableNamesFromDB"

   ' Message frequency "Exporting rec x of y"
   Public Const MSG_FREQUENCY As Int32 = 100

   ' Default SELECT statement
   Public Const SQL_SELECT As String = "*"

#End Region

   Sub Main(ByVal cmdLineArgs() As String)

      ' ToDo: XML configuration for "free form SQL" export, allowing JOINs etc.

      AppIntro("MsSqlDataToText")
      AppCopyright()
      BlankLine()

      Console.WriteLine("Using configuration file: {0}", cmdLineArgs(0))
      BlankLine()

      ' *** Debug
      ' AnyKey()

      ' Parse the configuration file
      Dim xmlConfig As New XmlConfig(cmdLineArgs(0))

      Dim eRet As XmlConfig.eCfgXMLResult = CType(xmlConfig.Init(), XmlConfig.eCfgXMLResult)

      Select Case eRet

         Case libXmlConfig.Tools.Xml.Configuration.XmlConfig.eCfgXMLResult.xmlErrFileNotFound
            Console.WriteLine("Configuration XML {0} not found.", cmdLineArgs(0))

         Case libXmlConfig.Tools.Xml.Configuration.XmlConfig.eCfgXMLResult.xmlErrNoConfigurationData
            Console.WriteLine("No valid configuration found in {0}.", cmdLineArgs(0))

         Case libXmlConfig.Tools.Xml.Configuration.XmlConfig.eCfgXMLResult.xmlErrOtherUnknown
            Console.WriteLine("Couldn't parse configuration XML {0}.", cmdLineArgs(0))

         Case libXmlConfig.Tools.Xml.Configuration.XmlConfig.eCfgXMLResult.xmlSuccess

            Console.WriteLine("Total records exported: {0}", ExportAllData(xmlConfig).ToString)

      End Select

   End Sub

   ''' <summary>
   ''' Main processing loop. Iterates through all tables in a database and exports the data 
   ''' of these tables to a CSV per table, according to the settings in the XML configuration file.
   ''' </summary>
   ''' <param name="xmlConfig">
   ''' Initialized XML configuration handler.
   ''' </param>
   ''' <returns>
   ''' Total number of records exported.
   ''' </returns>
   Function ExportAllData(ByVal xmlConfig As XmlConfig) As Int32

      Dim sConnectionstring As String = String.Empty
      sConnectionstring = xmlConfig.DefaultConfig.Sections.GetSection(CFG_DATABASE).Entrys.GetEntry("ConnectionString").Value.ToString

      Dim sTableQuery As String = String.Empty
      sTableQuery = xmlConfig.DefaultConfig.Sections.GetSection(CFG_DATABASE).Entrys.GetEntry("TableQuery").Value.ToString

      ' Retrieve all tables from the database schema
      Dim dbCon As New SqlConnection(sConnectionstring)
      Dim dbCmd As New SqlCommand(sTableQuery, dbCon)

      ' Load up the tables and schemta in a dataset
      Dim dbAdapter As New SqlDataAdapter(dbCmd)
      Dim dbDataset As New DataSet

      dbAdapter.Fill(dbDataset)

      ' Section for storing table names
      ' AddTableNamesXML
      Dim bolAddTableNames As Boolean
      Dim xmlSection As IXmlCfgSection = xmlConfig.DefaultConfig.Sections.GetSection(CFG_EXPORT)
      If xmlSection.Entrys.HasEntry("AddTableNamesXML") AndAlso CType(xmlSection.Entrys.GetEntry("AddTableNamesXML").Value, Boolean) Then
         bolAddTableNames = True
         If Not xmlConfig.DefaultConfig.Sections.HasSection(CFG_TABELNAMESFROMDB) Then
            xmlConfig.DefaultConfig.Sections.CfgSections.Add(New XmlCfgSection(CFG_TABELNAMESFROMDB))
         End If
      End If


      ' Loop through all tables and export a CSV of the Table Data
      Dim lRet, lTotal As Int32
      For Each row As DataRow In dbDataset.Tables(0).Rows

         ' Add to the XML, might be useful for SkipColumns
         If bolAddTableNames = True Then
            xmlSection = xmlConfig.DefaultConfig.Sections.GetSection(CFG_TABELNAMESFROMDB)
            Dim xmlEntry As New XmlCfgEntry(CFG_TABELNAMESFROMDB, "ColumnName", String.Format("{0}.{1}", row(0), row(1)))
            xmlSection.Entrys.CfgEntrys.Add(xmlEntry)
         End If

         lRet = ExportTableData(xmlConfig.DefaultConfig, dbCon, row)
         lTotal += lRet

         Console.WriteLine(" - Records exported: " & lRet.ToString)
         BlankLine()

      Next

      ' Only save configuration, we changed it
      If bolAddTableNames = True Then
         xmlConfig.SaveConfig()
      End If

      If dbCon.State = ConnectionState.Closed Then
         dbCon.Open()
      End If

      Return lTotal

   End Function

   Function ExportTableData(ByVal xmlCfg As IXmlCfg, ByVal dbCon As SqlConnection, ByVal row As DataRow) As Int32
      ' Export the data of a specific table

      ' Create the text file name, based upon the table's name
      Dim sOutFile As String = String.Empty
      sOutFile = CreateFilename(xmlCfg, row)
      Console.WriteLine("-- Export file: " & sOutFile)

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

         ' Should not take longer than 2 minutes.
         dbCmd.CommandTimeout = 240

         ' *** Count the number of rows in that table
         Dim dbReader As SqlDataReader = Nothing
         Try

            dbReader = dbCmd.ExecuteReader()

            Do While dbReader.Read
               lRows = libBAUtil.DBTools.SQL.SqlDBUtil.GetDBInteger(dbReader, "RecCount")
            Loop


         Catch exsql As SqlException

            BlankLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("A database error while exporting data:")
            BlankLine()
            Console.ForegroundColor = ConsoleColor.Gray
            Console.WriteLine("Trying to export table {0}, continuing with next table.", sTablename)
            BlankLine()
            Console.ForegroundColor = ConsoleColor.DarkGray
            Console.WriteLine(exsql.Message)
            BlankLine()

         Catch ex As Exception

            BlankLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("An error while exporting data:")
            BlankLine()
            Console.ForegroundColor = ConsoleColor.Gray
            Console.WriteLine("Trying to export table {0}, continuing with next table.", sTablename)
            BlankLine()
            Console.ForegroundColor = ConsoleColor.DarkGray
            Console.WriteLine(ex.Message)
            BlankLine()

         Finally

            Console.ForegroundColor = ConsoleColor.Gray
            If Not dbReader Is Nothing Then
               dbReader.Close()
            End If

         End Try

         ' Skip empty table altogether?
         If lRows = 0 Then
            If (xmlCfg.Sections.GetSection(CFG_EXPORT).Entrys.HasEntry("SkipEmptyTables") = True) AndAlso
               (CBool(xmlCfg.Sections.GetSection(CFG_EXPORT).Entrys.GetEntry("SkipEmptyTables").Value) = True) Then
               Console.WriteLine(" - Skipping empty table: " & sTablename)
               Return lRows
            End If
         End If


         ' *** Do the actual data export of the table
         ' Retrieve the column names
         Dim sqlSelect As String = GetSqlSelectClause(xmlCfg, sTablename)
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
         Dim sDelim As String = xmlCfg.Sections.GetSection(CFG_EXPORT).Entrys.GetEntry("ColumnDelimiter").Value.ToString

         sQuery = System.String.Format("SELECT" & sqlSelect & "FROM [{0}].[{1}]", row(0), row(1))
         dbCmd.CommandText = sQuery

         Console.WriteLine(" - Using: {0}", sQuery)

         Try

            dbReader = dbCmd.ExecuteReader()

            Dim lRecCount As Int32

            Do While dbReader.Read

               Dim txtLine As String = String.Empty

               For i As Int32 = 0 To dbReader.FieldCount - 1

                  ' Skip column?
                  If Not DoSkipColumn(xmlCfg, sTablename, dbReader.GetName(i)) Then

                     Dim sTemp As String = String.Empty

                     With dbReader

                        ' ** BUNAC special **
                        ' Integers (row IDs) stored in binary columns. Convert data from those to a hex string
                        If Not .IsDBNull(i) And (.GetFieldType(i) Is GetType(Byte())) Then ' .DataType.ToString = "System.Byte[]"
                           Dim abyt() As Byte = CType(.GetValue(i), Byte())
                           If abyt.Length <= 16 Then
                              For Each b As Byte In abyt
                                 sTemp &= Convert.ToString(b, 16)
                              Next
                              sTemp = "0x" & sTemp
                           Else
                              For Each b As Byte In abyt
                                 sTemp &= Convert.ToString(b)
                              Next
                           End If
                        Else
                           sTemp = ReadNullAsEmptyString(dbReader, i)
                        End If

                        sTemp = SanitizeCSV(sTemp, sDelim)

                        If txtLine.Length > 0 Then
                           txtLine &= sDelim
                        End If

                        txtLine &= EnQuote(sTemp)

                     End With

                  End If

               Next

               TxtWriteLine(sOutFile, txtLine)

               lRecCount += 1
               If lRecCount Mod MSG_FREQUENCY = 0 Then
                  Console.SetCursorPosition(0, Console.CursorTop)
                  Console.Write(" - Exporting record {0} of {1}", lRecCount.ToString, lRows.ToString)
               End If

            Loop

            ' Fake the last update
            Console.SetCursorPosition(0, Console.CursorTop)
            Console.WriteLine(" - Exporting record {0} of {1}", lRows.ToString, lRows.ToString)
            'BlankLine()

         Catch exsql As SqlException

            BlankLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("A database error while exporting data:")
            BlankLine()
            Console.ForegroundColor = ConsoleColor.Gray
            Console.WriteLine("Trying to export table {0}, continuing with next table.", sTablename)
            BlankLine()
            Console.ForegroundColor = ConsoleColor.DarkGray
            Console.WriteLine(exsql.Message)
            BlankLine()

         Catch ex As Exception

            BlankLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("An error while exporting data:")
            BlankLine()
            Console.ForegroundColor = ConsoleColor.Gray
            Console.WriteLine("Trying to export table {0}, continuing with next table.", sTablename)
            BlankLine()
            Console.ForegroundColor = ConsoleColor.DarkGray
            Console.WriteLine(ex.Message)
            BlankLine()

         Finally

            Console.ForegroundColor = ConsoleColor.Gray
            dbReader.Close()

         End Try

      End If

      Return lRows

   End Function

   Function CreateFilename(ByVal xmlCfg As IXmlCfg, ByVal row As DataRow, Optional ByVal sExtension As String = ".csv") As String

      ' Construct the output file name
      Dim sRet As String = String.Empty

      sRet = NormalizePath(xmlCfg.Sections.GetSection(CFG_EXPORT).Entrys.GetEntry("DestinationPath").Value.ToString) & row(0).ToString & "_" & row(1).ToString & sExtension
      Return sRet

   End Function

   Function CreateTextFileHeader(xmlCfg As IXmlCfg, ByVal exportFile As String, ByVal tableName As String, ByVal dbCols As DataColumnCollection) As Boolean

      ' Safe guard - write header at all?
      If xmlCfg.Sections.GetSection(CFG_EXPORT).Entrys.HasEntry("ColumnNameAsFirstLine") Then
         Dim o As IXmlCfgEntry
         o = xmlCfg.Sections.GetSection(CFG_EXPORT).Entrys.GetEntry("ColumnNameAsFirstLine")
         If CBool(o.Value) = False Then
            Return True
         End If
      End If

      Dim sDelim As String = String.Empty
      Dim sHdrText As String = String.Empty

      sDelim = xmlCfg.Sections.GetSection(CFG_EXPORT).Entrys.GetEntry("ColumnDelimiter").Value.ToString

      For Each dbCol As DataColumn In dbCols

         If Not DoSkipColumn(xmlCfg, tableName, dbCol.ColumnName) Then
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

   ''' <summary>
   ''' Determine if a database table should be at least partially exported at all.
   ''' </summary>
   ''' <param name="xmlCfg">
   ''' Current configuration object.
   ''' </param>
   ''' <param name="tableName">
   ''' Check this table. tableName is in the format (schema).(database), e.g. dob.CustomerData
   ''' </param>
   ''' <returns>
   ''' Export this table? <see langword="true"/>/<see langword="false"/>.
   ''' </returns>
   Function DoSkipTable(ByVal xmlCfg As IXmlCfg, ByVal tableName As String) As Boolean

      ' Determine if a table should be skipped
      ' Section ExportColumns takes precedence over section SkipColumns, so check that first
      If xmlCfg.Sections.HasSection(CFG_EXPORTCOLUMNS) = True Then

         ' If the table name is mentioned anywhere, it shouldn't be skipped completely
         ' To prevent similar table names resulting in the irregular creation of empty CSV files, add
         ' the '.' to the comparison.
         ' Fixes issue https://dev.sta.net/Knuth.Konrad/mssqldatatotext/issues/3
         For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection(CFG_EXPORTCOLUMNS).Entrys.CfgEntrys
            If Left(o.Value.ToString, tableName.Length + 1) = tableName & "." Then
               Return False
            End If
         Next

         Return True

      Else

         For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection(CFG_SKIPCOLUMNS).Entrys.CfgEntrys
            If o.Value.ToString = tableName & ".*" Then
               Return True
            End If
         Next

         Return False

      End If

   End Function

   Function DoSkipColumn(ByVal xmlCfg As IXmlCfg, ByVal tableName As String, ByVal columnName As String) As Boolean

      ' Entry in XML is of format tableName.columnName or tableName.*

      ' Section ExportColumns takes precedence over section SkipColumns, so check that first
      If xmlCfg.Sections.HasSection(CFG_EXPORTCOLUMNS) = True Then

         For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection(CFG_EXPORTCOLUMNS).Entrys.CfgEntrys
            If o.Value.ToString = String.Format("{0}.{1}", tableName, columnName) OrElse o.Value.ToString = tableName & ".*" Then
               Return False
            End If
         Next

         Return True

      Else

         ' Determine if a certain column should be skipped
         For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection(CFG_SKIPCOLUMNS).Entrys.CfgEntrys
            If o.Value.ToString = String.Format("{0}.{1}", tableName, columnName) Then
               Return True
            End If
         Next

         Return False

      End If

   End Function

   Function GetSqlSelectClause(ByVal xmlCfg As IXmlCfg, ByVal tableName As String) As String
      ' Retrieve the SQL SELECT clause for the specified table

      ' Default hard-coded value
      Dim sResult As String = SQL_SELECT

      ' Default value from configuration, if any
      For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection(CFG_TABLESELECT).Entrys.CfgEntrys
         If Left(o.Value.ToString, 2) = "*|" Then
            sResult = Mid(o.Value.ToString, 3)
         End If
      Next

      ' Specific select for this table?
      For Each o As IXmlCfgEntry In xmlCfg.Sections.GetSection(CFG_TABLESELECT).Entrys.CfgEntrys
         If Left(o.Value.ToString, tableName.Length + 1) = tableName & "|" Then
            sResult = Mid(o.Value.ToString, tableName.Length + 2)
         End If
      Next

      ' Console.WriteLine("SELECT from configuration: {0}", sResult)

      Return " " & sResult & " "

   End Function

   Function SanitizeCSV(ByVal csvData As String, ByVal colDelim As String, Optional ByVal colDelimSubstitute As String = "|") As String
      ' Sanitize CSV data

      Dim sResult As String = csvData

      If sResult.Contains(vbQuote()) Then
         ' Single double quote with two double quotes
         sResult = sResult.Replace(vbQuote(), vbQuote(2))
      End If

      ' Column delimiter within data?
      sResult = sResult.Replace(colDelim, colDelimSubstitute)

      Return sResult

   End Function

End Module
