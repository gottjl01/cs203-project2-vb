Namespace PFW.CSIST203.Project2.Persisters.Excel

    ''' <summary>
    ''' Generic interface used for persisting data to and from a data source
    ''' </summary>
    Public Interface IPersistData
        Inherits IDisposable

        ''' <summary>
        ''' Retrieves a specific row number from the data source using a unique ID
        ''' </summary>
        ''' <param name="id">The unique identifier used by this persister to retrieve specific rows</param>
        ''' <returns>The data row representing the requested data</returns>
        Function GetRow(id As Integer) As System.Data.DataRow

        ''' <summary>
        ''' Retrieves a count of the number of elements present in the data source
        ''' </summary>
        ''' <returns>The number of items present in the underlying data source</returns>
        Function CountRows() As Integer


    End Interface

End Namespace


