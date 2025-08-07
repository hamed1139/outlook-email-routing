Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    RouteNewMails EntryIDCollection
End Sub

Private Sub RouteNewMails(ByVal EntryIDCollection As String)
    Dim ns As Outlook.NameSpace
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlSheet As Object
    Dim path As String
    Dim arr() As String
    Dim entryID As Variant
    Dim obj As Object
    Dim itm As Outlook.MailItem
    Dim subjectText As String
    Dim lastRow As Long
    Dim i As Long
    Dim keyword1 As String, keyword2 As String, keyword3 As String
    Dim folderName As String
    Dim targetFolder As Outlook.folder

    ' Path to the Excel file containing routing rules
    path = "C:\Users\mahmoudiniah\Desktop\Email_Routing_Rules.xlsx"

    Set ns = Application.GetNamespace("MAPI")
    arr = Split(EntryIDCollection, ",")

    ' Open Excel - create a new instance to avoid interfering with other open workbooks
    Set xlApp = CreateObject("Excel.Application")
    If xlApp Is Nothing Then Exit Sub

    Set xlWb = xlApp.Workbooks.Open(path, ReadOnly:=True)
    Set xlSheet = xlWb.Sheets(1)

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' xlUp

    ' Process each email
    For Each entryID In arr
        Set obj = ns.GetItemFromID(entryID)
        If TypeName(obj) = "MailItem" Then
            Set itm = obj
            subjectText = UCase(itm.Subject)

            ' Compare with keywords
            For i = 2 To lastRow
                keyword1 = UCase(Trim(xlSheet.Cells(i, 1).Value))
                keyword2 = UCase(Trim(xlSheet.Cells(i, 2).Value))
                keyword3 = UCase(Trim(xlSheet.Cells(i, 3).Value))
                folderName = Trim(xlSheet.Cells(i, 4).Value)

                If InStr(subjectText, keyword1) > 0 And InStr(subjectText, keyword2) > 0 And InStr(subjectText, keyword3) > 0 Then
                    Set targetFolder = FindFolderByName(folderName)
                    If Not targetFolder Is Nothing Then
                        itm.Move targetFolder
                        Exit For
                    End If
                End If
            Next i
        End If
    Next entryID

    ' Close only the file and app instance opened by this macro
    xlWb.Close False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
End Sub

' Search for folder by name recursively
Private Function FindFolderByName(folderName As String) As Outlook.folder
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim rootFolder As Outlook.folder
    Dim target As Outlook.folder

    For Each rootFolder In ns.Folders
        Set target = SearchFoldersRecursively(rootFolder, folderName)
        If Not target Is Nothing Then
            Set FindFolderByName = target
            Exit Function
        End If
    Next rootFolder
End Function

Private Function SearchFoldersRecursively(startFolder As Outlook.folder, folderName As String) As Outlook.folder
    If UCase(startFolder.Name) = UCase(folderName) Then
        Set SearchFoldersRecursively = startFolder
        Exit Function
    End If

    Dim subFolder As Outlook.folder
    For Each subFolder In startFolder.Folders
        Set SearchFoldersRecursively = SearchFoldersRecursively(subFolder, folderName)
        If Not SearchFoldersRecursively Is Nothing Then Exit Function
    Next subFolder
End Function
