Namespace PFW.CSIST203.Project2

    Public Class Util

        ''' <summary>
        ''' Returns an excel file connection string suitable for use by an OleDbConnection
        ''' </summary>
        ''' <param name="excelFile">Path or filename of an excel document on disk</param>
        ''' <returns>A connection string that is suitable for selecting all non-header content from the excel file</returns>
        Public Shared Function GetExcelConnectionString(excelFile As String, hasHeaderRow As Boolean) As String

            ' retrieve the extension and initialize connection string builder
            Dim extension = System.IO.Path.GetExtension(excelFile)
            Dim builder As New System.Data.OleDb.OleDbConnectionStringBuilder()
            Dim header As String = IIf(hasHeaderRow, "Yes", "No").ToString()

            ' if we are using Office 2000-era excel files, use the 4.0 provider
            If String.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) AndAlso Not System.Environment.Is64BitOperatingSystem Then
                builder.Provider = "Microsoft.Jet.OLEDB.4.0"
                builder.Add("Extended Properties", String.Format("Excel 8.0;IMEX=1;HDR={0};", header))
                ' if we are dealing with Office 2007+ Excel, use the 12.0 provider
            ElseIf String.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) Then
                builder.Provider = "Microsoft.ACE.OLEDB.12.0"
                builder.Add("Extended Properties", String.Format("Excel 12.0;IMEX=1;HDR={0};", header))
            Else
                ' The provider cannot be determined and an exception must be thrown
                Throw New NotSupportedException(String.Format("Excel connection string for files with extension '{0}' are not supported by the operating system", extension))
            End If
            builder.DataSource = excelFile
            Return builder.ConnectionString

        End Function

    End Class

End Namespace


