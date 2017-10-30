Imports OfficeOpenXml
Imports System.IO

Public NotInheritable Class EPPLibrary

    Private Shared fileTarget As String
    Private Shared excelApplication As ExcelPackage
    Private Shared workbook As ExcelWorkbook = Nothing
    Private Shared activeWorkSheet As ExcelWorksheet = Nothing

    ''Private Shared excelApplication As New Excel.Application
    ''Private Shared workbook As Excel.Workbook = Nothing
    ''Private Shared activeWorkSheet As Excel._Worksheet = Nothing

    Public Shared Sub setFile(ByVal fileName As String)
        fileTarget = fileName
        excelApplication = New ExcelPackage(New FileInfo(fileTarget))
        workbook = excelApplication.Workbook
    End Sub

    Public Shared Sub setWorkSheet(ByVal index As Integer)
        activeWorkSheet = workbook.Worksheets(index)
    End Sub

    Public Shared Function getValueFrom(ByVal x As Integer, ByVal y As Integer)
        If activeWorkSheet Is Nothing Then
            Throw New Exception("No Worksheet selected error. Use the setWorkSheet() method first from ExcelLibrary.vb")
        Else
            Dim range As ExcelRange = activeWorkSheet.Cells(x, y)
            Return range.Value.ToString()
            ''Dim range As Excel.Range = activeWorkSheet.UsedRange
            ''Return range.Cells(x, y).Value2.ToString()
        End If
    End Function

    Public Shared Function getWorkSheets()
        Dim list As New List(Of ComboboxItem)
        For i As Integer = 1 To workbook.Worksheets.Count
            Dim sheet As ExcelWorksheet
            sheet = workbook.Worksheets(i)
            Dim cbItem As New ComboboxItem
            cbItem.Text = sheet.Name
            cbItem.Value = i
            list.Add(cbItem)
        Next
        Return list
    End Function


End Class
