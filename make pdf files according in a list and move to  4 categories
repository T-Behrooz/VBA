Sub Macro1()

Dim folderPath, myfolderName, subfolder, sep, filename_1, copyPath As String
Dim moin(2 To 5) As String
sep = "\"
myfolderName = "performance"
filename_1 = "_Úãá˜ÑÏ"
folderPath = "d:\" & myfolderName & "_" & Format(Date, "yyyy-mm-dd") '-- main folders name

'--- create main folder ------------------------------------------------------------------------
If Len(Dir(folderPath, vbDirectory)) = 0 Then
    MkDir (folderPath)
End If
'--- create sub folders ------------------------------------------------------------------------
For i = 2 To 5
subfolder = folderPath & "\" & Sheet2.Range("k" & i).Value '-- subFolders acopyPathress in excel sheets
moin(i) = Sheet2.Range("k" & i).Value  ' store categories in array
 If Len(Dir(subfolder, vbDirectory)) = 0 Then
      MkDir (subfolder)
 End If
Next i
'-------------------------------------------------------------------------------------------------
Sheet3.Activate    ' activate main sheet
 For i = 2 To 39
    Range("c1").Value = Sheet2.Range("b" & i).Value  ' store c1 cell by given data
    Range("c1").Select
    For j = 2 To 5    ' making pdf files
     If Sheet3.Range("i1").Text = moin(j) Then
     copyPath = folderPath & "\" & moin(j) & "\" & Range("c1").Text & filename_1 & ".pdf"
     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
     copyPath, Quality:=xlQualityStandard, IncludeDocProperties:= _
     True, IgnorePrintAreas:=False, OpenAfterPublish:=False
     End If
    Next j
 Next i
MsgBox " All Files  in " & folderPath & " Copied!"


End Sub
