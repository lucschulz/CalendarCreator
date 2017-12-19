Attribute VB_Name = "ExportAllModules"
Sub ExportVBA()
    Dim path As String
        path = "C:\Users\luc\OneDrive\Code\GitHub\CalendarCreator\source\"
        
            If Dir(path, vbDirectory) = "" Then
                MkDir (path)
            End If
    
    Dim i As Integer
        For i = 1 To ActiveWorkbook.VBProject.VBComponents.Count
            Debug.Print ActiveWorkbook.VBProject.VBComponents(i).Name
            ActiveWorkbook.VBProject.VBComponents(i).Export (path & "\" & ActiveWorkbook.VBProject.VBComponents(i).Name & ".bas")
        Next i
    
    'Call DeleteUnnecessaryFiles(path)
        
End Sub

Sub DeleteUnnecessaryFiles(path As String)
    Dim fileToDelete As Variant
        If Right$(path, 1) <> "\" Then path = path & "\"
                
    fileToDelete = FileList(path, "*.frx")
        Dim i As Long
            For i = LBound(fileToDelete) To UBound(fileToDelete)
                Kill path & fileToDelete(i)
            Next
End Sub

Function FileList(folder As String, Optional fileExtension As String = "*.*") As Variant

    Dim sTemp As String, sHldr As String
    
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
        sTemp = Dir(folder & fileExtension)
    If sTemp = "" Then
        FileList = Split("No files found", "|")  'ensures an  array is returned
        Exit Function
    End If
    
    Do
        sHldr = Dir
        If sHldr = "" Then Exit Do
        sTemp = sTemp & "|" & sHldr
     Loop
    FileList = Split(sTemp, "|")
End Function

Sub OpenOutputFolder(rootPath As String)
    Dim processID As Double
        Const fileToLaunch = "explorer.exe"
            Const arguments = " "
                processID = Shell(fileToLaunch & arguments & rootPath, 3)
End Sub

