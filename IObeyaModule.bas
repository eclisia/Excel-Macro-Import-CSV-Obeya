Attribute VB_Name = "IObeyaModule"
Sub OpenAndParse_ObeyaCSV()
'   Author : FTA
'   Description :
'   This macro permits to import the CSV file from the IObeya application
'   After imported, it parse the file.
'   Then, the column name are setted as the filter for the column
'
'   Version 00    Creation of the macro

On Error GoTo GestionError
    'Variable definition
    Dim myWorkbook As Workbook
    Dim myWorkbooks As Workbooks
    Dim myWorkSheetSommaire As Worksheet
    Dim myWorkSheetImport As Worksheet
    Dim myFilePath As String
    Dim myDate As Date
    Dim myRangeTitle As Range

'    myFilePath = "D:\Users\taintuf\Desktop\CCSL_Contr_le___Bancs-OBEYA_HREO-19122016-153049.csv"
    
    'Get the current workbooks
    Set myWorkbooks = Application.Workbooks
    
    'Boucle de test - A supprimer car ne sert à rien
    For Each myWorkbook In myWorkbooks
        Debug.Print "Name of the current Workbook : " & myWorkbook.Name
    Next myWorkbook
    
    'Get the current workbook
    Set myWorkbook = Application.ActiveWorkbook
    Debug.Print "Name of the current Workbook : " & myWorkbook.Name
    
    Set myWorkSheetSommaire = myWorkbook.Worksheets.Item("Sommaire")
    Debug.Print "Name of the current Worksheet : " & myWorkSheetSommaire.Name
    'Get
    myFilePath = myWorkSheetSommaire.Range("B7").Value
    Debug.Print "Nom du fichier contenu dans la case B7 : " & myFilePath
    
    If (FileExiste(myFilePath) = False) Then
        MsgBox "Fichier CSV introuvable ou cellule 'B7' vide", vbCritical, "Erreur Fichier CSV"
        Exit Sub
    End If
    
    
    If (CheckExtension(myFilePath) = False) Then
        Exit Sub
    End If
    
    
    'Create new Worksheet
    myWorkbook.Worksheets.Add After:=myWorkbook.Worksheets.Item("Sommaire")
    Debug.Print "Name of the created Worksheet : " & myWorkbook.ActiveSheet.Name
    Set myWorkSheetImport = myWorkbook.ActiveSheet
    myWorkSheetImport.Name = "ExportIObeya_" & RecuperationtDate
    Debug.Print "New Name of the created Worksheet : " & myWorkSheetImport.Name
    
    

'    Call the opentext function on the workbookS
    With myWorkSheetImport.QueryTables.Add(Connection:="TEXT;" & myFilePath, Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    'Rename Column
    '1 - 8
    Range("A1").Value = "Description Action"
    Range("B1").Value = "Projet"
    Range("C1").Value = "Porteur"
    Range("D1").Value = "Week"
    Range("G1").Value = "Type"
    Range("H1").Value = "Sous-Type"
  
    'Delete unused column
    'E - F - I - J
    Range("E:E,F:F,I:I,J:J").Delete
    
    
    Set myRangeTitle = Range("A1:Z1")
    myRangeTitle.AutoFilter
    

    'Format the Title row
    With Range("A1:M1")
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .EntireRow.AutoFit
    End With

    MsgBox "Import CSV Obeya Terminé", vbInformation, "Import"
    
    Exit Sub
    
    
GestionError:
    Select Case Err.Number
        Case 1004
            MsgBox "Erreur, la feuille existe déjà avec ce même nom", vbCritical, "Erreur 1004 - Nom de Feuille"
            Exit Sub
    End Select
    Resume

End Sub


Function RecuperationtDate() As String
'This is a Generic function to get Date
    Dim myDate As Date
    myDate = Date
    RecuperationtDate = Format(myDate, "yyyy-MM-dd")
    
    Debug.Print "mydate " & myDate
    
    
End Function
Private Function FileExiste(ByVal path As String) As Boolean
'This function verify the presence of the given file
    If (Len(Dir(path)) > 0) Then
        FileExiste = True
    End If
    
End Function

Private Function CheckExtension(ByVal path As String) As Boolean
    Dim strext As String
    
    'Set the default true value
    CheckExtension = True
    With CreateObject("Scripting.FileSystemObject")
        strext = .GetExtensionName(path)    'ActiveWorkbook.path)
    End With
    
    Debug.Print "Extension du fichier est : " & strext
    
    If StrComp(strext, "csv", vbBinaryCompare) <> 0 Then
        MsgBox "Tentative d'ouverture d'un fichier autre que *.csv", vbCritical, "Erreur de type de fichier"
        'replace the boolean value
        CheckExtension = False
    End If
    
    
End Function

Public Sub LoadNewObeyaCSVFile()
    Dim myFileName As Variant
    Dim myFilter As String
    myFilter = "Fichier CSV (*.csv),*.csv"
    
    myFileName = Application.GetOpenFilename(FileFilter:=myFilter, FilterIndex:=1, Title:="Selectionner le fichier CSV de l'Obeya", MultiSelect:=False)
    Debug.Print myFileName
    
    Application.ActiveSheet.Range("B7").Value = myFileName
    

End Sub
