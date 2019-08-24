Namespace PFW.CSIST203.Project2.Persisters.Excel

    ''' <summary>
    ''' Excel Persister that interacts with data in an xls or xlsx file
    ''' </summary>
    Public Class ExcelPersister
        Implements IPersistData

        Private logger As log4net.ILog = log4net.LogManager.GetLogger(GetType(ExcelPersister))

        Private _Data As System.Data.DataTable = Nothing

        ''' <summary>
        ''' This data table must be populated with all data contained in the specified excel file
        ''' </summary>
        Friend Property Data As System.Data.DataTable
            Get
                Return _Data
            End Get
            Private Set(value As System.Data.DataTable)
                _Data = value
            End Set
        End Property

        Private _isDisposed As Boolean = False

        ''' <summary>
        ''' Get a value indicating whether or not the object has been disposed
        ''' </summary>
        Friend Property isDisposed As Boolean
            Get
                Return _isDisposed
            End Get
            Private Set(value As Boolean)
                _isDisposed = value
            End Set
        End Property


        ''' <summary>
        ''' This contructor creates a persister that contains no data
        ''' </summary>
        Public Sub New()
            Data = New DataTable("Sheet1")
            Data.Columns.AddRange(
                {
                    New DataColumn("First Name", GetType(String)),
                    New DataColumn("Last Name", GetType(String)),
                    New DataColumn("E-mail Address", GetType(String)),
                    New DataColumn("Business Phone", GetType(String)),
                    New DataColumn("Company", GetType(String)),
                    New DataColumn("Job Title", GetType(String))
                })
        End Sub

        Public Sub New(excelFilepath As String)
            Throw New NotImplementedException()
        End Sub

        Public Function GetRow(rowNumber As Integer) As System.Data.DataRow Implements IPersistData.GetRow
            Throw New NotImplementedException()
        End Function

        Public Function CountRows() As Integer Implements IPersistData.CountRows
            Throw New NotImplementedException()
        End Function

        Public Sub Dispose() Implements IPersistData.Dispose
            ' TODO: Implement this method
        End Sub

    End Class

End Namespace


