Attribute VB_Name = "Module3"
Public CECE_file
Public Opcion_T1XX
Public Cr_number
Public UltimoRenglon
Public Tipo
Public Proliferacion
Public IndexDeCECE
Public IndexDeSD
Public IndexDeST
Public Column_Ident
Public Column_Formato
Public Bandera_Activar_Cobre
Public Name_Tab
Public FlagNewColumns
Public PartNumberColumn_3
'Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public TipoOrig


Sub Open_file(Tipo) 'step 1 & 2

    Dim linea As String
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Bandera_Activar_Cobre = Empty
    
    
    If BatchProcesing = Empty Then
        Module3.OpenExcelFile_CECE CECE_file, EndFlag       'Nombre del archivo CECE
        On Error Resume Next
        Workbooks(CECE_file).Sheets("CECE").AutoFilterMode = False
            If EndFlag = "Terminar" Then Exit Sub
    End If
    
    'check for tab
    'add sheet if nesesary
    tab1 = "Pivot Summary"
    ChecarWSht tab1, CECE_file

    
    
    
    If Tipo = "PRICE DETAIL - Impl In Amort" Then
        TestAmort Flag26, Hoja26, AmortPriceRow, AmortPriceCol    'Checa si archivo tiene la tabla de Amortizacion 2/26/20
        If Flag26 = Empty And BatchProcesing = Empty Then
            Respuesta = MsgBox("Warning!!, no Amortization information found in CECE file, Do you want to add the table?", vbOKCancel)
            If Respuesta = 1 Then
                'PorcentajePen.PorcenPenetration
                AmortTable.Show
                TestAmort Flag26, Hoja26, AmortPriceRow, AmortPriceCol
            End If
            
            If Respuesta = 2 Then
                End
            End If
        End If
    End If
    
    
    If Flag26 = "Ok" Or Tipo = "PRICE DETAIL - Impl In Burden" Then   'Si el CECE tiene el archivo de cost procede a procesar
    
        Checkbox11 = ThisWorkbook.Worksheets("Sheet1").Shapes("Check Box 11").ControlFormat.Value     'checar precios de price history DB
        If Checkbox11 = 1 Then
            CECEvsCOst.CECEvsCost1
        End If
        
        ComponentSupplierName CECE_file, FlagNewColumns, PartNumberColumn_3         'Se movio por que se requeria
        If FlagNewColumns = "FormatoNuevo" Then
        
        Else
            detector_of_GM_File EndFlag
            If EndFlag = "Terminar" Then Exit Sub
        End If
        
        If Proliferacion = "Si" Then 'Agregado 8/15/18
            GeneradorDeSubfijos
        End If
        
        Advertencia_num_partes_iguales CECE_file, EndFlag     'Advertencia de num de arnes repetidos en el CECE
            If EndFlag = "Terminar" Then Exit Sub
       
        If Workbooks(CECE_file).Sheets("CECE").Range("C29") = Empty Then
            MsgBox ("CECE format is empty")
            Exit Sub
        End If
        
        'StartTime3 = Timer
        'Advertencia_costo_Cobre_diferentes EndFlag
            If EndFlag = "Terminar" Then Exit Sub
        'SecondsElapsed3 = Round(Timer - StartTime3, 2)
        
        'Advertencia_costo_Cobre_diferentes_rapida EndFlag
         '   If EndFlag = "Terminar" Then Exit Sub
        Checkbox8 = ThisWorkbook.Worksheets("Sheet1").Shapes("Check Box 8").ControlFormat.Value
        If BatchProcesing = Empty And Checkbox8 = 1 Then
            Detector_De_Cobre_Mal EndFlag
                If EndFlag = "Terminar" Then Exit Sub
                
            Alarmas.DetectorDeEngCostExtra EndFlag
                If EndFlag = "Terminar" Then Exit Sub
                
            If Workbooks(CECE_file).Sheets("CECE").Range("C29") = Empty Then
                MsgBox ("CECE format is empty")
                Exit Sub
            End If
        End If
        
        Application.ScreenUpdating = True
        Application.StatusBar = "Preparing file for processing"
        DoEvents
        Application.ScreenUpdating = False
        
        BorradorDeRenglonesVacios
        borraTapeStart
        
        'Workbooks(CECE_file).Sheets("CECE").Rows(29).Copy
        'Workbooks(CECE_file).Sheets("CECE").Rows(29).PasteSpecial Paste:=xlPasteValues
        
        
        'PartNumberColumn_3 Sheet6
        LetraN = PartNumberColumn_3 - 15
        ColumnLetter_1 = Split(Cells(1, LetraN).Address, "$")(1)
        ColumnLetter_2 = Split(Cells(1, PartNumberColumn_3).Address, "$")(1)
        
        'Workbooks(CECE_file).Sheets("CECE").Range("DH29:DZ29").Copy
        
        RenglonFinal = 29
        Do
            CeldaF = Workbooks(CECE_file).Sheets("CECE").Cells(RenglonFinal, 3).Value
            RenglonFinal = RenglonFinal + 1
        Loop Until CeldaF = Empty
        
        
        
        Workbooks(CECE_file).Sheets("CECE").Range(ColumnLetter_1 & 29 & ":" & ColumnLetter_2 & RenglonFinal).Copy
        Workbooks(CECE_file).Sheets("CECE").Range(ColumnLetter_1 & 29 & ":" & ColumnLetter_2 & RenglonFinal).PasteSpecial Paste:=xlPasteValues
        
        
        
        
        NoCostCr
        If Proliferacion = "Si" Then 'Agregado 8/15/18
        
        Else
            limpiar_numeros_de_parte_espacios CECE_file
        End If
        'BorradorDeRenglonesVacios
        CopiarHarness
        Workbooks(CECE_file).SaveAs , FileFormat:=xlOpenXMLWorkbookMacroEnabled
        CECE_file = ActiveWorkbook.Name
            
        'Workbooks(CECE_file).Close
        'Workbooks.Open CECE_file, UpdateLinks:=False
        
        Workbooks(CECE_file).Activate
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        Delete_Tab_Cr_Number CECE_file
        Sort
        
        If Opcion_T1XX = "Si" Then
            'Application.EnableEvents = False
            'ActiveSheet.DisplayPageBreaks = False
            Crear_T1XX_Format
            'Module3.T1XX_Clasifier
            Clasifier.T1XX_faster_Clasifier
            Module3.T1XX_Cleanser
            Module3.Deleter_T1XX_Tab
        End If
        
        copiar_numeros_de_parte 'Copia y borra numeros de parte secundarios
        
        IndexDeCECE = Workbooks(CECE_file).Sheets("CECE").Index
        IndexDeSD = Workbooks(CECE_file).Sheets("Summary Data").Index
        IndexDeST = Workbooks(CECE_file).Sheets("Summary Temp").Index
        Workbooks(CECE_file).Sheets("CECE").Activate
        
        Workbooks(CECE_file).Sheets("Summary Data").Unprotect "PASSWORD"        '10/17/19
        Application.ScreenUpdating = True
        Application.StatusBar = "Pivot Summary"
        DoEvents
        Application.ScreenUpdating = False
        Module9.PivotSummary
        
        Application.ScreenUpdating = True
        DoEvents
        Application.StatusBar = "Preparing GM Detail"
        DoEvents
        Application.ScreenUpdating = False
        
        Paste_numeros_de_parte  'Pastea numeros de parte al CECE
        
         '1/5/17
        'If Opcion_T1XX = "Si" Then
        '    Module3.T1XX_Cleanser
        'End If
        
        'Application.Calculation = xlCalculationAutomatic
        
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(Tipo).Copy _
            after:=Workbooks(CECE_file).Sheets("CECE")                  'step 7     'Copia el tab en question
        
        Workbooks(CECE_file).Sheets("Summary Table").Visible = True     'step 8 & 9
        
        Id21 = Workbooks(CECE_file).Sheets("CECE").Range("C28").Value
        If Id21 = "Aptiv P/N" Then
            Workbooks(CECE_file).Sheets("Summary Table").Unprotect Password:="PASSWORD"
        Else
            Workbooks(CECE_file).Sheets("Summary Table").Unprotect Password:="password"
        End If
    
        ThisWorkbook.Sheets("Summary Table").Range("O2:R3").Copy Destination:=Workbooks(CECE_file).Sheets("Summary Table").Range("O2")      'step 10 & 11
        
        'Modulo para contar renglones de Summary Table
        Summary_Table_Row = 4
        Do
                Summary_Table_Row = Summary_Table_Row + 1
        Loop Until IsEmpty(Workbooks(CECE_file).Sheets("Summary Table").Cells(Summary_Table_Row, 1))
        Summary_Table_Row = Summary_Table_Row - 1
    
        platform = Workbooks(CECE_file).Sheets("CECE").Cells(20, 3).Text   'step 12
        Workbooks(CECE_file).Sheets("Summary Table").Rows("3:3").AutoFilter
        Workbooks(CECE_file).Sheets("Summary Table").Range("$A$3:$R$" & Summary_Table_Row).AutoFilter field:=2, Criteria1:= _
        platform
        
        ThisWorkbook.Sheets("Summary Table").Range("O21:R21").Copy
        Workbooks(CECE_file).Sheets("Summary Table").Range("O4:R" & Summary_Table_Row).SpecialCells(xlCellTypeVisible).PasteSpecial xlPasteFormulasAndNumberFormats
        
        'Seccion del T1XXX agregada en 4/28/16 pero desabilitada por falta de aprovacion
        Set c = ThisWorkbook.Sheets("Sheet2").Columns(4).Find(platform, LookIn:=xlValues)
        If Not c Is Nothing Then Workbooks(CECE_file).Sheets("Summary Table").Range("Q2").Value = Application.WorksheetFunction.VLookup(platform, ThisWorkbook.Worksheets("Sheet2").Range("D1:E100"), 2, False)
                
        On Error Resume Next
        linea = Workbooks(CECE_file).Name
        linea2 = Workbooks(CECE_file).LinkSources(xlExcelLinks)
        If Not IsEmpty(linea2) Then
        linea3 = 1
            Do
                Workbooks(CECE_file).ChangeLink Name:=linea2(linea3), NewName:=linea, Type:=xlExcelLinks
                linea3 = linea3 + 1
            Loop Until linea3 = UBound(linea2)
        End If
        
        If BatchProcesing = Empty Then
            With Workbooks(CECE_file).Sheets(Tipo).Range("A1:O21")
                Set rFind = .Find(what:=".xlsm", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, searchformat:=False)
                If Not rFind Is Nothing Then
                    Encon = rFind.Row
                    'b = rFind.Column
                End If
            End With
    
            'Remueve path de formulas
            If Encon > 1 Then
                Dim rCell As Range
                For Each rCell In Workbooks(CECE_file).Sheets(Tipo).UsedRange
                    rCell.Replace what:="[*]", replacement:=""
                Next
            End If
            Workbooks(CECE_file).Sheets("PRICE DETAIL - Impl In Burden ").Range("N4").Formula = "=OFFSET('Pivot Summary'!D$4,0,((COLUMN(A1)-1)/-2))"
    
            Valor = Workbooks(CECE_file).Names
        
        
        End If
                        
'        With Workbooks(CECE_file).Sheets("PRICE DETAIL - Impl In Burden").Range("A1:O21")
'            Set rFind = .Find(what:=".xlsm", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, searchformat:=False)
'            If Not rFind Is Nothing Then
'                Encon = rFind.Row
'                'b = rFind.Column
'            End If
'        End With
'
'        'Remueve path de formulas
'        If Encon > 1 Then
'            Dim rCell As Range
'            For Each rCell In Workbooks(CECE_file).Sheets("PRICE DETAIL - Impl In Burden").UsedRange
'                rCell.Replace what:="[*]", replacement:=""
'            Next
'        End If
'        Workbooks(CECE_file).Sheets("PRICE DETAIL - Impl In Burden ").Range("N4").Formula = "=OFFSET('Pivot Summary'!D$4,0,((COLUMN(A1)-1)/-2))"
'
        'Valor = Workbooks(CECE_file).Names
        On Error Resume Next
        Dim nm As Name
        For Each nm In Workbooks(CECE_file).Names
            nm.Delete
        Next nm
              
        part_number_column CECE_file, PartNumberColumn_2                    'Step 14
        Unit_Copper_weight_lbsColumn = PartNumberColumn_2 + 14 - 1
        PartNumberColumn = Application.WorksheetFunction.VLookup(PartNumberColumn_2, ThisWorkbook.Worksheets("Sheet4").Range("A9:B414"), 2)
        Unit_Copper_weight_lbsColumn = Application.WorksheetFunction.VLookup(Unit_Copper_weight_lbsColumn, ThisWorkbook.Worksheets("Sheet4").Range("A9:B414"), 2)
                    
        'stept 16 & 17
        Workbooks(CECE_file).Sheets("CECE").Range(PartNumberColumn & ":" & Unit_Copper_weight_lbsColumn).Name = "CompInfo"
        
        copiar_num_partes (CECE_file)           'Esto es para copiar num de parte orignal a nuevos
        'Step 18
        
        'Raw_link = Workbooks(CECE_file).Sheets(Tipo).Cells(18, 4).Formula
        'Link_1 = Left(Raw_link, Len(Raw_link) - 17)
        'Link_to_be_used = Right(Link_1, Len(Link_1) - 2)
        
        Workbooks(CECE_file).Sheets(Tipo).Range("A3:O43").Select
        Selection.Replace what:="[CECE]", replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False
        
        Selection.Replace what:="[Pivot Summary]", replacement:="", _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
            False, searchformat:=False, ReplaceFormat:=False
            
        'step 20/21/22
        Count_insertrow_scrolldown CECE_file, Tipo
        'Corrector de formulas de Kilos
        Convertidor_de_Kilos CECE_file, Tipo
        
        
        'Format unit copper lbs cell bold red if value is 0 for cable, terminals, or misc sub-assemblies
        'Dice que termina pero todavia falta mucho
        'Module5.Terminar CECE_file, Tipo, Notice_flag_1
            
        'Step 32
        Tier2PN_ComponentSupplierName_ComponentOrigin
        Application.Calculation = xlCalculationAutomatic
        'Module3.Copiador_De_Cobre  'Agregado 2/1/17 para comprobacion
        'Workbooks(CECE_file).Sheets(Tipo).Copy , After:=Sheets("CECE")
        Cr_number = Right(Workbooks(CECE_file).Worksheets("CECE").Cells(10, 3).Value, 6)
        'ActiveSheet.Name = "GM Detail " & Cr_number
        'ActiveSheet.Name = "CR" & Cr_number & "-GM Detail"
        Name_Tab = "CR" & Cr_number & "-GM Detail"      'Added on 4/11/16
        'ActiveSheet.Name = Name_Tab
        
        'Agrega Investment Amortization
        Inv_Amort AmortPriceRow, AmortPriceCol
        
        Set ws = Sheets(Tipo)
        ws.Copy after:=Sheets("CECE")
        Set wsnew = Sheets(Sheets("CECE").Index + 1)
        wsnew.Name = Name_Tab
        
        Workbooks(CECE_file).Worksheets("T1XX_Format").Name = "CR " & Cr_number & " 200mm & Gage Rule"
        
        'If Not Bandera_Activar_Cobre = "No" Then
        '    Module3.Copiador_De_Cobre  'Actualizado en 3/27/17
            
        'End If
        
        'Advertencias que no agregan valor
        'Prelim_pricer.Advertencia_GM_Detail_Descr_Repetidos
        'Prelim_pricer.Advertencia_faltan_nombre_de_Cliente
        Module1.CleansePriceDetail
        
        'b = "GM Detail " & Cr_number
        'Copiador_Maestro.Copiador_Maestro b, CECE_file           'Esta rutina salva info a archivo maestro
        
        'Module3.Numero_de_parte_del_cliente Cr_number
        Main.Main
        'Application.Calculation = xlCalculationAutomatic
        'Module3.Redondear 'disabled per Roger York 4/13/17
        'Module3.Copiador_De_Cobre  'Agregado 2/1/17 para comprobacion
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Versione = Left(Right(ThisWorkbook.Name, 10), 6)
        
        
        Workbooks(CECE_file).Sheets(Name_Tab).Cells(22, 1) = Versione
        Workbooks(CECE_file).Sheets(Name_Tab).Cells(22, 1).Font.Color = vbWhite
        Workbooks(CECE_file).Sheets(Name_Tab).Cells(11, 3).Select
        
    
        If Not Bandera_Activar_Cobre = "No" Then
            Module3.CopiadorMejor  'Actualizado en 2/16/17
            
        End If
        Application.ScreenUpdating = True
        
        
        If Proliferacion = "Si" Then 'Agregado 9/19/18
            BorradorDeSubfijos
        End If
        
        
        Workbooks(CECE_file).Save
        Comunicacion.Comunicacion1
        Module4.Create_Log_De_CR_2                'Added on 5/11/17
        
        'If BatchProcesing = Empty Then
            'Module5.Terminar CECE_file, Tipo, Notice_flag_1
        'End If
        
        
        Checkbox10 = ThisWorkbook.Worksheets("Sheet1").Shapes("Check Box 10").ControlFormat.Value
            If Checkbox10 = 1 Then
                    
                Application.StatusBar = "Checking network..."
                    'Workbooks.Open fileName:="http://p04.na.delphiauto.net/16/101/dpnasm/sales/am/gmam/Shared%20Documents/Standard%20Pricing%20Documents/GM%20ECQ/GM%20SA.xlsb", ReadOnly:=True
                If Err.Number = 1004 Then
                    'MsgBox ("There is not Delphi Network Available, please connect to access GM SA File")
                    'EndFlag = "Terminar"
                Else
                    cmdUploadPrice_Click_T1XX Tipo
                    
                End If
                
            Else
                'Module5.Terminar CECE_file, Tipo, Notice_flag_1
            End If
        
    Else
        Workbooks(CECE_file).Close
    End If
    
    Dump.Dump1
    
    Application.ScreenUpdating = False
    Checkbox17 = ThisWorkbook.Worksheets("Sheet1").Shapes("Check Box 17").ControlFormat.Value
    If BatchProcesing = Empty And Checkbox17 = 1 Then
        ReportForIndCr '10/02/20
    End If
    
    If BatchProcesing = Empty Then  'Copiado en 5/27/19 por que no estaba poniendo terminado al final
        Module5.Terminar CECE_file, Tipo, Notice_flag_1
    End If
    
    Application.DisplayAlerts = False        'Added on 4/10/17
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
EndNow:
    a = Err.Description
    If Err.Description = "Subscript out of range" Then
        Workbooks(wb).Sheets.Add.Name = "Pivot Summary"
    End If
    
    
    
End Sub


Sub OpenExcelFile_CECE(CECE_file, EndFlag)  'Pide archivos para abrir

     Dim vFile As Variant
      
     CECE_file = Application.GetOpenFilename("Excel Files (*.xl*)," & _
    "*.xl*", 1, "Open CECE", "Open", False)
     'If Cancel then exit

     If TypeName(CECE_file) = "Boolean" Then
        EndFlag = "Terminar"
        Exit Sub
     End If

     If CECE_file = Empty Or CECE_file = False Then
                    EndFlag = "Terminar"
                    Exit Sub
     End If
     
     Workbooks.Open CECE_file, UpdateLinks:=False ', ReadOnly:=True     was causing problems with xlsm files detected by RY
     FullPath = CECE_file
     FunctionGetFileName FullPath, GetFileName
     
     CECE_file = GetFileName
 End Sub

Public Function FunctionGetFileName(FullPath, GetFileName)

Dim StrFind As String
    Do Until Left(StrFind, 1) = "\"
        iCount = iCount + 1
        StrFind = Right(FullPath, iCount)
            If iCount = Len(FullPath) Then Exit Do
    Loop

    GetFileName = Right(StrFind, Len(StrFind) - 1)

End Function

Sub Sort() 'step 3
'
' Sort Macro
'
    CurrentCol = 16
    Do
            CurrentCol = CurrentCol + 1
    Loop Until IsEmpty(Cells(20, CurrentCol + 1))
'

    Last_LetterColumn = Application.WorksheetFunction.VLookup(CurrentCol, ThisWorkbook.Worksheets("Sheet4").Range("A9:B414"), 2)
    Columns("Q:" & Last_LetterColumn).Select
    ActiveWorkbook.Worksheets("CECE").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CECE").Sort.SortFields.Add Key:=Range("Q20:" & Last_LetterColumn & "20"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("CECE").Sort
        .SetRange Range("Q1:" & Last_LetterColumn & "2000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Public Function Last_part_number_column(PartNumberColumn)                     ' Subrutina cuenta la cantidad de columnas oara orice
    
    SalesRow = 28
    SalesColumn = 20
      
    Do
        SalesCellValue = Cells(SalesRow, SalesColumn).Value
        If Not SalesCellValue = "Sales Input Price  $" Then
            SalesColumn = SalesColumn + 1
        End If
    
    Loop Until SalesCellValue = "Sales Input Price  $"


End Function

Public Function part_number_column(CECE_file, PartNumberColumn_2)                    ' Subrutina cuenta la cantidad de columnas oara orice
    
    SalesRow = 28
    PartNumberColumn_2 = 20
      
    Do
        SalesCellValue = Workbooks(CECE_file).Sheets("CECE").Cells(SalesRow, PartNumberColumn_2).Value
        If Not SalesCellValue = "Part Number" Then
            PartNumberColumn_2 = PartNumberColumn_2 + 1
        End If
    
    Loop Until SalesCellValue = "Part Number"


End Function

Sub Count_insertrow_scrolldown(CECE_file, Tipo)
    Rows_by_default = 20
    pivot_row = 5
    Additional_rows = 0
    UltimoRenglonFormatoCE = 29
    'step 20
    Do
            If Not Workbooks(CECE_file).Sheets("Pivot Summary").Cells(pivot_row, 1) = "(blank)" Then
                pivot_row = pivot_row + 1
            End If
            a = Workbooks(CECE_file).Sheets("Pivot Summary").Cells(pivot_row, 1)
    Loop Until IsEmpty(Workbooks(CECE_file).Sheets("Pivot Summary").Cells(pivot_row, 3)) Or Workbooks(CECE_file).Sheets("Pivot Summary").Cells(pivot_row, 1) = "(blank)"
    
    'Step 21
    If pivot_row - 3 > Rows_by_default Then
        Additional_rows = pivot_row - Rows_by_default
        Do
            Workbooks(CECE_file).Sheets(Tipo).Range("B24").EntireRow.Offset(1, 0).Insert
            Additional_rows = Additional_rows - 1
        Loop Until Additional_rows = 0
    End If
    
    'Step 22
    If pivot_row > 6 Then
        'Workbooks(CECE_file).Sheets(Tipo).Range("B24:O24").Select
        'Selection.AutoFill Destination:=Range("B24:O" & pivot_row + 24 - 6), Type:=xlFillDefault
        
        PivotRow = 6
        GMDetail_Row = 25
        Do
            Workbooks(CECE_file).Sheets(Tipo).Cells(GMDetail_Row, 2) = Workbooks(CECE_file).Sheets("Pivot Summary").Cells(PivotRow, 2).Value
            If Workbooks(CECE_file).Sheets("Pivot Summary").Cells(PivotRow, 1) = Empty Then
                Workbooks(CECE_file).Sheets(Tipo).Cells(GMDetail_Row, 6) = Workbooks(CECE_file).Sheets("Pivot Summary").Cells(PivotRow - 1, 1).Value
            Else
                Workbooks(CECE_file).Sheets(Tipo).Cells(GMDetail_Row, 6) = Workbooks(CECE_file).Sheets("Pivot Summary").Cells(PivotRow, 1).Value
            End If
            
            Workbooks(CECE_file).Sheets(Tipo).Cells(GMDetail_Row, 6).NumberFormat = "@"
            Valor_numerico = Workbooks(CECE_file).Sheets(Tipo).Cells(GMDetail_Row, 6).Value
            Texto = Application.WorksheetFunction.Text(Valor_numerico, 0)
            Workbooks(CECE_file).Sheets(Tipo).Cells(GMDetail_Row, 6) = "Test"
            Workbooks(CECE_file).Sheets(Tipo).Cells(GMDetail_Row, 6) = Texto
            
            PivotRow = PivotRow + 1
            GMDetail_Row = GMDetail_Row + 1
        Loop Until Workbooks(CECE_file).Sheets("Pivot Summary").Cells(PivotRow, 2) = "(blank)"
                
        Workbooks(CECE_file).Sheets(Tipo).Range("G24:O24").Select
        Selection.AutoFill Destination:=Range("G24:O" & pivot_row + 24 - 6), Type:=xlFillDefault
    End If
    
    'Step 24
        Module4.Copy_above CECE_file, Tipo
    
    
    'Step 25
    Alt_CECE = CECE_file
    Delete_row CECE_file, Tipo

    
    
        columns_pivot = 4
        Do
            If Not Workbooks(CECE_file).Sheets("Pivot Summary").Cells(4, columns_pivot) = "(blank)" Then
                columns_pivot = columns_pivot + 1
            End If
            a = Workbooks(CECE_file).Sheets("Pivot Summary").Cells(4, columns_pivot)
        Loop Until IsEmpty(Workbooks(CECE_file).Sheets("Pivot Summary").Cells(4, columns_pivot)) Or Workbooks(CECE_file).Sheets("Pivot Summary").Cells(4, columns_pivot) = "(blank)"
    
    'Step 27
        OrgColumn = 14
        Ciclos = columns_pivot - 4
        Do
            If Ciclos > 1 Then
                OrgColumn = OrgColumn + 2
                Workbooks(CECE_file).Sheets(Tipo).Columns(14).Copy Destination:=Workbooks(CECE_file).Sheets(Tipo).Columns(OrgColumn)
                Workbooks(CECE_file).Sheets(Tipo).Columns(15).Copy Destination:=Workbooks(CECE_file).Sheets(Tipo).Columns(OrgColumn + 1)
                Ciclos = Ciclos - 1
                
            End If
        Loop Until Ciclos = 1 Or Ciclos = 0 'trampa para cuando no detecta numeros de parte para los ciclos 6/19/14
        
    'Step 29
        columns_Price_Detail = 8
        CurrentCol_CECE = 16

        Do  'determina cual es la ultima columna del Price Detail
            columns_Price_Detail = columns_Price_Detail + 1
        Loop Until IsEmpty(Workbooks(CECE_file).Sheets(Tipo).Cells(23, columns_Price_Detail + 1))
        
        Do
            UltimoRenglonFormatoCE = UltimoRenglonFormatoCE + 1
        Loop Until IsEmpty(Workbooks(CECE_file).Sheets("CECE").Cells(UltimoRenglonFormatoCE, 3))
        
        
        Do  'determina cual es la ultima columna del CECE
            CurrentCol_CECE = CurrentCol_CECE + 1
        Loop Until IsEmpty(Workbooks(CECE_file).Sheets("CECE").Cells(20, CurrentCol_CECE + 1))
        
        Last_LetterColumn_Price_Detail = Application.WorksheetFunction.VLookup(columns_Price_Detail, ThisWorkbook.Worksheets("Sheet4").Range("A9:B414"), 2)
        Last_LetterColumn_CECE = Application.WorksheetFunction.VLookup(CurrentCol_CECE, ThisWorkbook.Worksheets("Sheet4").Range("A9:B414"), 2)
        
                                                                                                
        Select Case Tipo
            Case Is = "PRICE DETAIL - Impl Separate"
                 Workbooks(CECE_file).Sheets("CECE").Range("Q7:" & Last_LetterColumn_CECE & "7").Formula = "=HLOOKUP(Q20,'PRICE DETAIL - Impl Separate'!$N4:$" & Last_LetterColumn_Price_Detail & "19,14,FALSE)"
            Case Is = "PRICE DETAIL - Impl In Burden"
                Workbooks(CECE_file).Sheets("CECE").Range("Q7:" & Last_LetterColumn_CECE & "7").Formula = "=HLOOKUP(Q20,'PRICE DETAIL - Impl In Burden'!$N4:$" & Last_LetterColumn_Price_Detail & "19,14,FALSE)"
            Case Is = "PRICE DETAIL-MCO"
                Workbooks(CECE_file).Sheets("CECE").Range("Q7:" & Last_LetterColumn_CECE & "7").Formula = "=HLOOKUP(Q20,'PRICE DETAIL-MCO'!$N4:$" & Last_LetterColumn_Price_Detail & "19,14,FALSE)"
        End Select
        
        
        
        'categorizador
        Renglonaevaluar = 24
        Do
            NumeroDelphi = Workbooks(CECE_file).Sheets(Tipo).Cells(Renglonaevaluar, 6)
            Descripcion_1 = Application.WorksheetFunction.VLookup(NumeroDelphi, Workbooks(CECE_file).Sheets("CECE").Range("C29:O" & UltimoRenglonFormatoCE), 13, False)
            Descripcion_2 = Application.WorksheetFunction.VLookup(NumeroDelphi, Workbooks(CECE_file).Sheets("CECE").Range("C29:O" & UltimoRenglonFormatoCE), 3, False)
            
            Renglon10 = 2
            Renglon11 = 2
            Do
                palabra = ThisWorkbook.Sheets("sheet2").Cells(Renglon10, 10).Text
                'palabra2 = ThisWorkbook.Sheets("sheet2").Cells(Renglon11, 11).Text
                If Descripcion_1 Like palabra Then
                    Renglon11 = Renglon10
                    Do
                        palabra2 = ThisWorkbook.Sheets("sheet2").Cells(Renglon11, 11).Text
                        If Descripcion_2 Like palabra2 Then
                            Bandera4 = "positivo"
                            Workbooks(CECE_file).Sheets(Tipo).Cells(Renglonaevaluar, 4) = ThisWorkbook.Sheets("sheet2").Cells(Renglon11, 12).Value
                        End If
                        
                        If palabra2 = "" Then
                            Workbooks(CECE_file).Sheets(Tipo).Cells(Renglonaevaluar, 4) = ThisWorkbook.Sheets("sheet2").Cells(Renglon11, 12).Value
                            Bandera4 = "positivo"
                        End If
                        
                        Renglon11 = Renglon11 + 1
                    Loop Until ThisWorkbook.Sheets("sheet2").Cells(Renglon11, 10) = Empty Or Bandera4 = "positivo"
                End If
                palabra = Empty
                Renglon10 = Renglon10 + 1
            Loop Until ThisWorkbook.Sheets("sheet2").Cells(Renglon10, 10) = Empty Or Bandera4 = "positivo"
            If Bandera4 = Empty Then
                Workbooks(CECE_file).Sheets(Tipo).Cells(Renglonaevaluar, 4) = "Other"
            End If
            Bandera4 = Empty
            Renglonaevaluar = Renglonaevaluar + 1
        Loop Until Workbooks(CECE_file).Sheets(Tipo).Cells(Renglonaevaluar, 6) = Empty
End Sub


Sub Delete_row(CECE_file, Tipo)
    'Step 25
    Delete_row_counter = 24
    On Error Resume Next
    Do
        Descrip_Price_detail_trunc = Workbooks(CECE_file).Sheets(Tipo).Cells(Delete_row_counter, 6).Value
        List_of_prohibited = Application.WorksheetFunction.VLookup(Descrip_Price_detail_trunc, ThisWorkbook.Worksheets("Sheet2").Range("B2:B100"), 1, False)
        If Not List_of_prohibited = Empty Then
            Workbooks(CECE_file).Sheets(Tipo).Rows(Delete_row_counter).EntireRow.Delete
        Else
            Delete_row_counter = Delete_row_counter + 1
        End If
        List_of_prohibited = Empty
        Descrip_Price_detail_trunc = Empty
    Loop Until Workbooks(CECE_file).Sheets(Tipo).Cells(Delete_row_counter, 13) = Empty

End Sub


Sub Convertidor_de_Kilos(CECE_file, Tipo)
    Renglon = 24
    
    Application.Calculation = xlAutomatic
    Final_Row_in_CECE = 29
        Do
            Final_Row_in_CECE = Final_Row_in_CECE + 1
        Loop Until IsEmpty(Workbooks(CECE_file).Sheets("CECE").Cells(Final_Row_in_CECE + 1, 3))
    
    Do
        If Workbooks(CECE_file).Sheets(Tipo).Cells(Renglon, 10).Text = "KG" Then
            Valor_de_Mspec = Workbooks(CECE_file).Sheets(Tipo).Cells(Renglon, 6).Text
            
                Lista_Mspec_EnCECE = Application.WorksheetFunction.VLookup(Valor_de_Mspec, Workbooks(CECE_file).Sheets("CECE").Range("C29:H" & Final_Row_in_CECE), 6, False)
                Lista_Mspec_EnCECE = UCase(Lista_Mspec_EnCECE)
                
                If Not Lista_Mspec_EnCECE = "KG" Then
                    Lista_Mspec = Application.WorksheetFunction.VLookup(Valor_de_Mspec, Workbooks(CECE_file).Sheets("Conv").Range("A2:J200"), 1, False)
                    
                    If Not Lista_Mspec = Empty Then
                        Workbooks(CECE_file).Sheets(Tipo).Range("L" & Renglon).Formula = "=(VLOOKUP(F" & Renglon & ",CompInfo,6,FALSE)*VLOOKUP(F" & Renglon & ",Conv!A1:J200,10,FALSE)*1000)-(K" & Renglon & "*$F$3)"
                        Workbooks(CECE_file).Sheets(Tipo).Range("J" & Renglon) = "M"
                                                                                                                    
                    End If
                End If
            
        End If
        
        'If Renglon = 53 Then Stop
        
        Lista_Mspec = Empty
        Renglon = Renglon + 1
        'If renglon = 36 Then Stop
    Loop Until Workbooks(CECE_file).Sheets(Tipo).Cells(Renglon, 6) = Empty
    Application.Calculation = xlCalculationManual
End Sub
            
Sub Delete_Tab_Cr_Number(CECE_file)
    Num_de_Cr = Empty
    Ident = Empty
    Num_Tab_Actual_de_Cr = 1
    EndFlag = Empty

    Do
        Num_de_Cr = Workbooks(CECE_file).Worksheets(Num_Tab_Actual_de_Cr).Cells(3, 5).Value
        Ident = Workbooks(CECE_file).Worksheets(Num_Tab_Actual_de_Cr).Cells(14, 2).Value
        If Num_de_Cr = "Copper Rate" And Ident = "CR Subject" Then
            Workbooks(CECE_file).Worksheets(Num_Tab_Actual_de_Cr).Delete
        Else
            Num_Tab_Actual_de_Cr = 1 + Num_Tab_Actual_de_Cr
        End If
    Conteo_Tab_Cons = Workbooks(CECE_file).Worksheets.Count
    Loop Until Num_Tab_Actual_de_Cr = Conteo_Tab_Cons + 1

End Sub
Sub copiar_num_partes(CECE_file)

    R2 = 24
    C2 = 17
    Do
        If Workbooks(CECE_file).Worksheets("CECE").Cells(R2, C2) = Empty Then
            Workbooks(CECE_file).Worksheets("CECE").Cells(R2, C2) = Workbooks(CECE_file).Worksheets("CECE").Cells(R2 + 2, C2).Value
        End If
        C2 = C2 + 1
    Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(26, C2) = Empty

End Sub


Sub limpiar_numeros_de_parte_espacios(CECE_file)
    Renglon_Num_parte = 20
    Columna = 17
    
    Do
        Do
            Num_de_Cr = Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, Columna).Value
            Num_de_Cr = Left(Num_de_Cr, 8)
            Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, Columna) = Num_de_Cr
            Columna = Columna + 1
        Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, Columna) = Empty
        Columna = 17
        Renglon_Num_parte = Renglon_Num_parte + 2
    Loop Until Renglon_Num_parte = 28
End Sub

Sub Numero_de_parte_del_cliente(Cr_number)
    
    Dim busqueda As String
    Dim Result As New MSXML2.DOMDocument
    Dim NumDeParteDelphi As String
    Dim list As IXMLDOMNodeList
    Dim attr As IXMLDOMAttribute
    Dim node As IXMLDOMNode
    Dim childNode As IXMLDOMNode
    Dim pctCompl As Single
     
    Renglon = 24
    'UltimoRenglon = Workbooks(CECE_file).Worksheets("GM Detail " & Cr_number).Cells(Rows.Count, "F").End(xlUp).Row
    UltimoRenglon = Workbooks(CECE_file).Worksheets(Name_Tab).Cells(Rows.Count, "F").End(xlUp).Row  'added on 4/11/17
    
    Do
        NumDeParteDelphi = Workbooks(CECE_file).Worksheets(Name_Tab).Cells(Renglon, 6).Value        'Added on 4/11/17
        'busqueda = "http://azurnaw-db025.northamerica.delphiauto.net/rWS/rWS_EXCEL/SALES/GetPartCrossRef/" & NumDeParteDelphi
        busqueda = "http://azurnaw-db025.aptiv.com/rWS/rWS_EXCEL/SALES/GetPartCrossRef/" & NumDeParteDelphi
        Set Service = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        Service.Open "GET", busqueda, False
        Service.send
         
        Result.async = False
        Result.LoadXML (Service.responseText)
                
        Set list = Result.SelectNodes("//GetPartCrossRefContract")
                
        For Each node In list
           Set attr = node.Attributes.getNamedItem("BusOrg_Id")
            If (Not attr Is Nothing) Then
                Debug.Print attr.BaseName & " " & attr.Text
            End If
            
            If obtenervalor(node, "Name") = "GENERAL MOTORS" Then
                NumGeneralMotors = obtenervalor(node, "Part_Nbr")
                If Len(NumGeneralMotors) = 8 Then
                    Workbooks(CECE_file).Worksheets(Name_Tab).Cells(Renglon, 5) = NumGeneralMotors      'Added on 4/11/17
                End If
            End If
            
            'If obtenervalor(node, "Supplier_Flag") = "Y" Then
            '    Workbooks(CECE_file).Worksheets("GM Detail " & Cr_number).Cells(Renglon, 8) = obtenervalor(node, "Name")
            'End If
                 

        Next node
        'Application.StatusBar = "Progress: " & x & " of 50: " & Format(x / 50, "0%")
        DoEvents
        
        Main.Main
        'Application.StatusBar = "Progress: " & Renglon - 24 & " of " & UltimoRenglon - 24 & " " & Format(Renglon / UltimoRenglon, "0%")
        
        'BarraDeProgreso.Show
        'pctCompl = Renglon
        'progress pctCompl
        
        Renglon = Renglon + 1
    Loop Until Renglon = UltimoRenglon + 1 Or Renglon > UltimoRenglon + 1
    
    Camino = ThisWorkbook.Path
    
    If ThisWorkbook.Sheets("Sheet1").Range("H23").Value = 1 Then
        'Application.Speech.Speak "I've finished. Where are you?"
        'PlaySound "C:\Sales\e-pricer\e-pricer\Sounds\Windows Notify.wav"
        'sndPlaySound32 "C:\Windows\Media\Windows Notify.wav", 0&
        sndPlaySound32 Camino & "\Windows Notify.wav", 0&
        'Camino & "\sounds\Windows Notify.wav"
        
    End If
      

    
    Application.StatusBar = False
End Sub


Function obtenervalor(node As IXMLDOMNode, xp As String)
    Dim n As IXMLDOMNode, nv
    Set n = node.SelectSingleNode(xp)
    If Not n Is Nothing Then obtenervalor = n.nodeTypedValue
End Function

Function Test_URLExists()
  Dim url As String
  
  'url = "http://www.mrexcel.com/forum/showthread.php?t=567315"
  'MsgBox url, vbInformation, URLExists(url)
  
  'url = "http://azurnaw-db025.northamerica.delphiauto.net/rWS/rWS_EXCEL/SALES/help/operations/Get_Part_Cross_Ref"
  url = "http://azurnaw-db025.aptiv.com/rWS/rWS_EXCEL/SALES/help/operations/Get_Part_Cross_Ref"
  MsgBox url, vbInformation, URLExists(url)
End Function

Function URLExists(url As String, rc) As Boolean
    Dim Request As Object
    Dim ff As Integer
    Dim rc As Variant
    
    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With Request
      .Open "GET", url, False
      .send
      rc = .statusText
    End With
    Set Request = Nothing
    'If rc = "OK" Then URLExists = True

    
    Exit Function
EndNow:
End Function



Sub progress(pctCompl As Single)

 UserForm1.Text.Caption = pctCompl & "% Completed"
 UserForm1.Bar.Width = pctCompl * 2

 DoEvents

End Sub

Sub Tier2PN_ComponentSupplierName_ComponentOrigin()
    If Not BatchProcesing = Empty Then
        'Tipo = TipoOrig
    End If
    Renglon = 24
    UltimoRenglon = Workbooks(CECE_file).Worksheets(Tipo).Cells(Rows.Count, "F").End(xlUp).Row
    Do
        'Tag = 0
        NumDeParteDelphi = Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 6).Value
        plant = Workbooks(CECE_file).Worksheets("CECE").Cells(20, 5).Value
        On Error Resume Next
        Err.Clear
        
        
        'DCS components
        CECE_row = Application.WorksheetFunction.Match(NumDeParteDelphi, Workbooks(CECE_file).Worksheets("CECE").Range("C1:C1000000"), 0)
        'Tag = 1
        If Workbooks(CECE_file).Worksheets("CECE").Cells(CECE_row, PartNumberColumn_2 + 7).Value = "DCS" Then
            Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 7) = Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 6).Value
            Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = "DCS"
            Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 9) = "NA"
            'Tag = 2
        End If
        
        
        'In-house cable
        If Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Empty Then   'Added on 4/26/16 Roger's
            CECE_row = Application.WorksheetFunction.Match(NumDeParteDelphi, Workbooks(CECE_file).Worksheets("CECE").Range("C1:C1000000"), 0)
                'Tag = 1
                If Workbooks(CECE_file).Worksheets("CECE").Cells(CECE_row, PartNumberColumn_2 + 11).Value = "M" And Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 4).Value = "Wire" Then
                    Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 7) = Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 6).Value
                    Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = "DELPHI"
                    Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 9) = "NA"
                    'Tag = 2
                End If
        End If
        
        'Lookup PN/Plant combination in GM SA file
        'This replaces the existing formulas in the GM detail, for more control
        'ComponentSupplierName CECE_file, FlagNewColumns, PartNumberColumn_3
        If FlagNewColumns = "FormatoNuevo" Then    'Esto decide de donde agarrar la informacion dependiendo del formato si es el nuevo o no
        
            If Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Empty Then   'Added on 4/26/16 Roger's
                'ColumnLetter1 = Split(Cells(1, PartNumberColumn_3).Address, "$")(1)
                'ColumnLetter2 = Split(Cells(1, PartNumberColumn_3 + 1).Address, "$")(1)
                'ColumnLetter3 = Split(Cells(1, PartNumberColumn_3 + 2).Address, "$")(1)
                'PartColumnCECE = Split(Cells(1, PartNumberColumn_3 - 15).Address, "$")(1)
                
                RowWDLocated = Application.WorksheetFunction.Match(NumDeParteDelphi, Workbooks(CECE_file).Worksheets("CECE").Range("C29:C5000"), 0) + 28
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 7) = Workbooks(CECE_file).Worksheets("CECE").Cells(RowWDLocated, PartNumberColumn_3)
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Workbooks(CECE_file).Worksheets("CECE").Cells(RowWDLocated, PartNumberColumn_3 + 1)
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 9) = Workbooks(CECE_file).Worksheets("CECE").Cells(RowWDLocated, PartNumberColumn_3 + 2)

            End If
        
      
        
        Else
        
            If Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Empty Then   'Added on 4/26/16 Roger's
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 7) = Application.WorksheetFunction.VLookup(NumDeParteDelphi & plant, Workbooks("GM SA.xlsb").Worksheets("DEEDS Plts").Range("C:N"), 9, False)
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Application.WorksheetFunction.VLookup(NumDeParteDelphi & plant, Workbooks("GM SA.xlsb").Worksheets("DEEDS Plts").Range("C:N"), 7, False)
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 9) = Application.WorksheetFunction.VLookup(NumDeParteDelphi & plant, Workbooks("GM SA.xlsb").Worksheets("DEEDS Plts").Range("C:N"), 12, False)
            End If
            
            'Lookup PN only in GM SA file
            If Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Empty Then
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 7) = Application.WorksheetFunction.VLookup(NumDeParteDelphi, Workbooks("GM SA.xlsb").Worksheets("DEEDS Plts").Range("A:N"), 11, False)
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Application.WorksheetFunction.VLookup(NumDeParteDelphi, Workbooks("GM SA.xlsb").Worksheets("DEEDS Plts").Range("A:N"), 9, False)
                Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 9) = Application.WorksheetFunction.VLookup(NumDeParteDelphi, Workbooks("GM SA.xlsb").Worksheets("DEEDS Plts").Range("A:N"), 14, False)
            End If
        End If
        
        'If still not found, leave blank
        If Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = Empty Then
            Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 7) = ""
            Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 8) = ""
            Workbooks(CECE_file).Worksheets(Tipo).Cells(Renglon, 9) = ""
        End If
                        
        Renglon = Renglon + 1
        
    Loop Until Renglon > UltimoRenglon Or Renglon > UltimoRenglon
    
    
End Sub

Sub detector_of_GM_File(EndFlag)

    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Flag = Workbooks("GM SA.xlsb").Worksheets("DEEDS Plts").Range("A5").Value 'Updated on 4/12/17
    'If Err.Number = 9 Then
    
    If Flag = Empty Then
        Application.StatusBar = "Please be patient while GM SA File loads..."
        
            'Workbooks.Open fileName:="http://p04.na.delphiauto.net/16/101/dpnasm/sales/am/gmam/Shared%20Documents/Standard%20Pricing%20Documents/GM%20ECQ/GM%20SA.xlsb", ReadOnly:=True
            Workbooks.Open fileName:="http://p04.na.aptiv.com/16/101/dpnasm/sales/am/gmam/Shared%20Documents/Standard%20Pricing%20Documents/GM%20ECQ/GM%20SA.xlsb", ReadOnly:=True
        
        If Err.Number = 1004 Then
            MsgBox ("There is not Delphi Network Available, please connect to access GM SA File")
            EndFlag = "Terminar"
        End If
    
    
    End If
    Application.StatusBar = ""
    Application.StatusBar = False
        'Resume Next
        
End Sub

Sub copiar_numeros_de_parte()

    Renglon_Num_parte = 29
    Renglon_temp = 2
    
    ThisWorkbook.Worksheets("Temp-1").Range("B2:B4140").Clear
    Do
        ThisWorkbook.Worksheets("Temp-1").Cells(Renglon_temp, 2).Value = Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, 4).Text
        Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, 4) = Empty
    Renglon_temp = Renglon_temp + 1
    Renglon_Num_parte = Renglon_Num_parte + 1
    Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, 2).Value = Empty

End Sub


Sub Paste_numeros_de_parte()

    Renglon_Num_parte = 29
    Renglon_temp = 2
    
    Do
        Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, 4) = ThisWorkbook.Worksheets("Temp-1").Cells(Renglon_temp, 2).Text
        
    Renglon_temp = Renglon_temp + 1
    Renglon_Num_parte = Renglon_Num_parte + 1
    Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon_Num_parte, 2).Value = Empty

End Sub

Sub Redondear()
    'UltimoRenglon
    Dim Renglon_GMDetail As Integer
    Renglon_GMDetail = 24
   
    TAB_1 = Name_Tab        'Added on 4/11/17
  
    'Do
    '    ValorDeCelda = Workbooks(CECE_file).Worksheets(TAB_1).Cells(Renglon_GMDetail, 12).Value
    '    ValorDeCelda = Round(ValorDeCelda, 4)
    '    Workbooks(CECE_file).Worksheets(TAB_1).Cells(Renglon_GMDetail, 12) = ValorDeCelda
    '    Renglon_GMDetail = Renglon_GMDetail + 1
    '    ValorDeCelda = Empty
    'Loop While Not Workbooks(CECE_file).Worksheets(TAB_1).Cells(Renglon_GMDetail, 2) = Empty
      
    
    Columna_Partes = 15
    Do
    Row_Precio = 6
        Do
            ValorDeCelda = Workbooks(CECE_file).Worksheets(TAB_1).Cells(Row_Precio, Columna_Partes).Value
            ValorDeCelda = Round(ValorDeCelda, 4)
            Workbooks(CECE_file).Worksheets(TAB_1).Cells(Row_Precio, Columna_Partes - 1) = ValorDeCelda
            Workbooks(CECE_file).Worksheets(TAB_1).Cells(Row_Precio, Columna_Partes) = ValorDeCelda
            ValorDeCelda = Empty
            Row_Precio = Row_Precio + 1
        Loop While Row_Precio < 20
        
        Columna_Partes = Columna_Partes + 2
    Loop While Not Workbooks(CECE_file).Worksheets(TAB_1).Cells(21, Columna_Partes - 1) = Empty
          
End Sub

Sub Advertencia_costo_Cobre_diferentes(EndFlag)
    
    Pivot_A = 29
    Pivot_B = 30

    part_number_column CECE_file, PartNumberColumn_2
    
    Do
        Do
            Numero_Pivot_A = Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_A, 3).Value
            Numero_Pivot_B = Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_B, 3).Value
            
            If Numero_Pivot_A = Numero_Pivot_B Then
                Cost_Pivot_A = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_A, PartNumberColumn_2 + 4).Value)
                Cost_Pivot_B = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_B, PartNumberColumn_2 + 4).Value)
                
                Copper_Pivot_A = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_A, PartNumberColumn_2 + 13).Value)
                Copper_Pivot_B = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_B, PartNumberColumn_2 + 13).Value)
                
                GAUGE_Pivot_A = Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_A, 16).Value
            
                Delta_Cost = Cost_Pivot_A - Cost_Pivot_B
                Delta_Copper = Copper_Pivot_A - Copper_Pivot_B
            
                'Cost
                If Delta_Cost <> 0 And GAUGE_Pivot_A = Empty Then
                
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_A, 3).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_B, 3).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_A, PartNumberColumn_2 + 4).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_B, PartNumberColumn_2 + 4).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                    EndFlag = "Terminar"
                    EndFlag_2 = "Terminar ciclo"
                End If
                    
                'Copper
                If Delta_Copper <> 0 And GAUGE_Pivot_A = Empty Then
                
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_A, 3).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_B, 3).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                                                
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_A, PartNumberColumn_2 + 13).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                                
                    With Workbooks(CECE_file).Sheets("CECE").Cells(Pivot_B, PartNumberColumn_2 + 13).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    EndFlag = "Terminar"
                    EndFlag_2 = "Terminar ciclo"
                End If
            
            
            End If
            Pivot_B = Pivot_B + 1
            Numero_Pivot_A = Empty
            Numero_Pivot_B = Empty
            Cost_Pivot_A = Empty
            Cost_Pivot_B = Empty
            GAUGE_Pivot_A = Empty
        Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_B, 3).Value = Empty
        EndFlag_2 = Empty
        Pivot_A = Pivot_A + 1
        Pivot_B = Pivot_A + 1
        Numero_Pivot_A = Empty
        Numero_Pivot_B = Empty
        Cost_Pivot_A = Empty
        Cost_Pivot_B = Empty
        GAUGE_Pivot_A = Empty
    Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(Pivot_A, 3).Value = Empty

    If EndFlag = "Terminar" Then
        AdventenciaCosto.Show
    End If
End Sub

Sub Borrar_TapeStart()
    
    Renglon = 29
    Do
        Part_Number = Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon, 3).Value
        If Part_Number = Empty Then
            Workbooks(CECE_file).Sheets("CECE").Row(Renglon).EntireRow.Delete
        End If
        Renglon = Renglon + 1
        
    Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon, 3).Value = Empty
    
End Sub


Sub T1XX_Cleanser()

    NombredeTab = "T1XX_Format"
    Row_1 = 8
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Cleaning Special format & CECE"
    Application.ScreenUpdating = False
    DoEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    
    Row_T1XX = 8
    Do
        
        If Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_T1XX, 1).Value = "NC" Then
            Counter_NC = Counter_NC + 1
        End If
        
        If Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_T1XX, 1).Value = "FC" Then
            Counter_FC = Counter_FC + 1
        End If
        
        Row_T1XX = Row_T1XX + 1
    
    Loop While Not Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_T1XX, 1).Value = Empty
    
    'BatchProcesing = "Yes" es para cuando corre el batch y no se requieren alarmas
    'If Not Counter_NC + 8 = Row_T1XX Then
        'Workbooks(CECE_file).Worksheets("T1XX_Format").Delete
        'Bandera_Activar_Cobre = "No"
        'Warning327.Show
    'End If

    
        'Do  'Esto se borra no importan las condiciones
            'If Workbooks(CECE_file).Worksheets(NombredeTab).Cells(Row_1, 3).Value = "TAPESTART" Or Workbooks(CECE_file).Worksheets(NombredeTab).Cells(Row_1, 3).Value = "STARTTAPE" Then
             '   Workbooks(CECE_file).Worksheets(NombredeTab).Rows(Row_1).EntireRow.Delete
             '   Workbooks(CECE_file).Worksheets("CECE").Rows(Row_1 + 21).EntireRow.Delete
         '   Else
         '       Row_1 = Row_1 + 1
         '   End If
         '   Valor = Empty
         '   Valor = Workbooks(CECE_file).Worksheets(NombredeTab).Cells(Row_1, 3).Value
            'If Row_1 > 2400 Then Stop
        'Loop While Not Valor = Empty
        
        
        'Experimento a ver si corre mas rapido
        
        'Workbooks(CECE_file).Worksheets("CECE").Cells.AutoFilter

        Dim Hoja As Worksheet
        'Dim hoja2 As Worksheet
        'Dim Info As Range
        'Dim Info2 As Range
        'Dim Info3 As Range
        'Application.DisplayAlerts = False
        
        Set Hoja = Workbooks(CECE_file).Worksheets("CECE")
        'With hoja
        '    Set Info = .Range("C28:C" & Row_T1XX + 20)
        'End With
        'On Error Resume Next
        'With Info
        '    .AutoFilter field:=1, Criteria1:="TAPESTART"
        '.Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Rows.Delete
        'End With
        
        'With hoja
        '    .AutoFilterMode = False
        '    If .FilterMode = True Then
        '        .ShowAllData
        '    End If
        'End With
        
        'Set hoja2 = Workbooks(CECE_file).Worksheets("T1XX_Format")
        'With hoja2
        '    Set Info2 = .Range("C8:C" & Row_T1XX)
        'End With
        
        'With Info2
        '    .AutoFilter field:=1, Criteria1:="TAPESTART"
        '.Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Rows.Delete
        'End With
        
        'With hoja2
        '    .AutoFilterMode = False
        '    If .FilterMode = True Then
        '        .ShowAllData
        '    End If
        'End With
        
        
    
        
  '  If Not Counter_NC + 8 = Row_T1XX Then
  '      Row_1 = 29
  '      Do
  '          If Workbooks(CECE_file).Worksheets("CECE").Cells(Row_1, 1).Value = "NC" Then
  '              Workbooks(CECE_file).Worksheets("CECE").Rows(Row_1).EntireRow.Delete
            
  '          Else
  '              Row_1 = Row_1 + 1
  '          End If
  '          Valor = Empty
  '          Valor = Workbooks(CECE_file).Worksheets("CECE").Cells(Row_1, 3).Value
            'If Row_1 > 2400 Then Stop
  '      Loop While Not Valor = Empty
        
  '  End If
        
        With Hoja
            .AutoFilterMode = False
            If .FilterMode = True Then
                .ShowAllData
            End If
        End With
    
        With Hoja
            Set Info3 = .Range("A28:A" & Row_T1XX + 20)
        End With
        On Error Resume Next
        With Info3
            .AutoFilter field:=1, Criteria1:="NC"
        .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Rows.Delete
        End With
        
        With Hoja
            .AutoFilterMode = False
            If .FilterMode = True Then
                .ShowAllData
            End If
        End With


        Row_1 = 29
        Do
            Workbooks(CECE_file).Worksheets("CECE").Cells(Row_1, 1) = ""
            Row_1 = Row_1 + 1
            Valor = Empty
            Valor = Workbooks(CECE_file).Worksheets("CECE").Cells(Row_1, 3).Value
            If Row_1 > 6000 Then Stop
        Loop While Not Valor = Empty
        
    Application.StatusBar = False
        
        
End Sub

Sub Crear_T1XX_Format()
    
    'Workbooks(CECE_file).Sheets.Add Before:=Worksheets(Worksheets.Count).Name = "T1XX_Format"
    Application.StatusBar = True
    DoEvents
    Application.ScreenUpdating = True
    Application.StatusBar = "Preparing Special format"
    Application.ScreenUpdating = False
    DoEvents
    Workbooks(CECE_file).Sheets.Add.Name = "T1XX_Format"
    CobreCECE = Workbooks(CECE_file).Worksheets("CECE").Range("L3").Value
    'ThisWorkbook.Worksheets("Machote").Range("A7:K7").Copy Destination:=Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A7")
    
    ThisWorkbook.Worksheets("Machote").Range("A7:K7").Copy
    With Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A7")
        .Cells(1).PasteSpecial xlPasteColumnWidths
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
    End With
    
    ThisWorkbook.Worksheets("Machote").Range("J1:K6").Copy
    With Workbooks(CECE_file).Worksheets("T1XX_Format").Range("J1")
        .Cells(1).PasteSpecial xlPasteColumnWidths
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
    End With
    
    Column_Ident = 12
    Do 'Part Number columns
        PartNumberColumn = Application.WorksheetFunction.VLookup(Column_Ident, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        PartNumberColumn_CECE = Application.WorksheetFunction.VLookup(Column_Ident + 5, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        ThisWorkbook.Worksheets("Machote").Range("L1:L7").Copy
        With Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & 1)
            .Cells(1).PasteSpecial xlPasteColumnWidths
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
            '.Cells(1).Select
            Application.CutCopyMode = False
        End With
        Workbooks(CECE_file).Worksheets("CECE").Range(PartNumberColumn_CECE & "19:" & PartNumberColumn_CECE & 22).Copy
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & 1).PasteSpecial xlPasteValues
        Workbooks(CECE_file).Worksheets("CECE").Range(PartNumberColumn_CECE & 24).Copy
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & 5).PasteSpecial xlPasteValues
        Workbooks(CECE_file).Worksheets("CECE").Range(PartNumberColumn_CECE & 26).Copy
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & 6).PasteSpecial xlPasteValues
        
        Column_Ident = Column_Ident + 1
    Loop Until Workbooks(CECE_file).Sheets("CECE").Cells(24, Column_Ident + 5).Value = Empty
    
    Ultimo_Renglon_CECE = Workbooks(CECE_file).Worksheets("CECE").Range("C29", Worksheets("CECE").Range("C28").End(xlDown)).Rows.Count
    'Ultimo_Renglon_CECE = 30
    'Driver
    PartNumberColumn = Application.WorksheetFunction.VLookup(Column_Ident, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
    ThisWorkbook.Worksheets("Machote").Range("M7").Copy
        With Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & 7)
            .Cells(1).PasteSpecial xlPasteColumnWidths
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
            '.Cells(1).Select
            Application.CutCopyMode = False
        End With
        
        'Parte baja de driver
        ThisWorkbook.Worksheets("Machote").Range("M20:M24").Copy
        With Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & Ultimo_Renglon_CECE + 2 + 8)
            .Cells(1).PasteSpecial xlPasteColumnWidths
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
            '.Cells(1).Select
            Application.CutCopyMode = False
        End With

    'Procedimiento que pone columnas de 3
    CECE_column = 17
    Column_Formato = Column_Ident + 1
    
    Do  'Esto pone cuadros parte de arriba
        ThisWorkbook.Worksheets("Machote").Range("N1:P7").Copy
        With Workbooks(CECE_file).Worksheets("T1XX_Format").Range(Cells(1, Column_Formato), Cells(7, Column_Formato + 3))
            .Cells(1).PasteSpecial xlPasteColumnWidths
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
            '.Cells(1).Select
            Application.CutCopyMode = False
        End With

        Dim ws3 As Worksheet, ws2 As Worksheet
        Set ws3 = Sheets("CECE")
        Set ws2 = Sheets("T1XX_Format")
        With ws3
            .Range(.Cells(19, CECE_column), .Cells(22, CECE_column)).Copy
        End With
        'With ws2
        '    .Paste Destination:=.Range(.Cells(1, Column_Formato), .Cells(1, Column_Formato + 2))
        'End With
        
        ws2.Cells(1, Column_Formato).PasteSpecial xlPasteValues
        ws2.Cells(1, Column_Formato + 1).PasteSpecial xlPasteValues
        ws2.Cells(1, Column_Formato + 2).PasteSpecial xlPasteValues
        
        'Range(Cells(2, Column_Formato), Cells(2, Column_Formato + 2)).Orientation = xlHorizontal
        'Range(Cells(4, Column_Formato), Cells(4, Column_Formato + 2)).Orientation = xlHorizontal
        'Range(Cells(2, Column_Formato), Cells(2, Column_Formato + 2)).Interior.ColorIndex = xlNone
        
        With Workbooks(CECE_file)
            Worksheets("T1XX_Format").Cells(5, Column_Formato) = Worksheets("CECE").Cells(24, CECE_column).Value
            Worksheets("T1XX_Format").Cells(5, Column_Formato + 1) = Worksheets("CECE").Cells(24, CECE_column).Value
            Worksheets("T1XX_Format").Cells(5, Column_Formato + 2) = Worksheets("CECE").Cells(24, CECE_column).Value
            Worksheets("T1XX_Format").Cells(6, Column_Formato) = Worksheets("CECE").Cells(26, CECE_column).Value
            Worksheets("T1XX_Format").Cells(6, Column_Formato + 1) = Worksheets("CECE").Cells(26, CECE_column).Value
            Worksheets("T1XX_Format").Cells(6, Column_Formato + 2) = Worksheets("CECE").Cells(26, CECE_column).Value
        End With
                        
        PartNumberColumn_Triple_A = Application.WorksheetFunction.VLookup(Column_Formato, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        PartNumberColumn_Triple_B = Application.WorksheetFunction.VLookup(Column_Formato + 2, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeTop).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeBottom).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeLeft).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeRight).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8).Borders(xlInsideHorizontal).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8).Borders(xlInsideVertical).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8).NumberFormat = "0.0000_) ;[red](0.0000)"
        
        'Parte baja 5 renglones
        UR5 = Ultimo_Renglon_CECE + 8 + 2
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & UR5 & ":" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6).Borders(xlEdgeTop).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & UR5 & ":" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6).Borders(xlEdgeBottom).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & UR5 & ":" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6).Borders(xlEdgeLeft).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & UR5 & ":" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6).Borders(xlEdgeRight).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & UR5 & ":" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6).Borders(xlInsideHorizontal).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & UR5 & ":" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6).Borders(xlInsideVertical).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn_Triple_A & UR5 & ":" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6).NumberFormat = "0.0000_) ;[red](0.0000)"
                                                                                                   
                                                                                                                                                                                                      '=O23+(P24-P23)*2.25
        PartNumberColumn_Triple_C = Application.WorksheetFunction.VLookup(Column_Formato + 1, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        UT1XX = Ultimo_Renglon_CECE + 8
        'ultimos 5 renglones de 3 formulas
        Range(Cells(UR5, Column_Formato), Cells(UR5, Column_Formato)).Formula = "=Sum(" & PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_A & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 1, Column_Formato), Cells(UR5 + 1, Column_Formato)).Formula = "=SUMIF($A8:A" & UT1XX & ",""VAVE""," & PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_A & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 2, Column_Formato), Cells(UR5 + 2, Column_Formato)).Formula = "=SUMIF($A8:A" & UT1XX & ",""NC""," & PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_A & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 3, Column_Formato), Cells(UR5 + 3, Column_Formato)).Formula = "=SUMIF($A8:A" & UT1XX & ",""FC""," & PartNumberColumn_Triple_A & "8:" & PartNumberColumn_Triple_A & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 4, Column_Formato), Cells(UR5 + 4, Column_Formato)).Formula = "=" & PartNumberColumn_Triple_A & Ultimo_Renglon_CECE + 8 + 5
        
        Range(Cells(UR5, Column_Formato + 1), Cells(UR5, Column_Formato + 1)).Formula = "=Sum(" & PartNumberColumn_Triple_C & "8:" & PartNumberColumn_Triple_C & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 1, Column_Formato + 1), Cells(UR5 + 1, Column_Formato + 1)).Formula = "=SUMIF($A8:A" & UT1XX & ",""VAVE""," & PartNumberColumn_Triple_C & "8:" & PartNumberColumn_Triple_C & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 2, Column_Formato + 1), Cells(UR5 + 2, Column_Formato + 1)).Formula = "=SUMIF($A8:A" & UT1XX & ",""NC""," & PartNumberColumn_Triple_C & "8:" & PartNumberColumn_Triple_C & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 3, Column_Formato + 1), Cells(UR5 + 3, Column_Formato + 1)).Formula = "=SUMIF($A8:A" & UT1XX & ",""FC""," & PartNumberColumn_Triple_C & "8:" & PartNumberColumn_Triple_C & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 4, Column_Formato + 1), Cells(UR5 + 4, Column_Formato + 1)).Formula = "=" & PartNumberColumn_Triple_C & Ultimo_Renglon_CECE + 8 + 5 & "+(" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 6 & "-" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 5 & ")*" & CobreCECE
        
        Range(Cells(UR5, Column_Formato + 2), Cells(UR5, Column_Formato + 2)).Formula = "=Sum(" & PartNumberColumn_Triple_B & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 1, Column_Formato + 2), Cells(UR5 + 1, Column_Formato + 2)).Formula = "=SUMIF($A8:A" & UT1XX & ",""VAVE""," & PartNumberColumn_Triple_B & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 2, Column_Formato + 2), Cells(UR5 + 2, Column_Formato + 2)).Formula = "=SUMIF($A8:A" & UT1XX & ",""NC""," & PartNumberColumn_Triple_B & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 3, Column_Formato + 2), Cells(UR5 + 3, Column_Formato + 2)).Formula = "=SUMIF($A8:A" & UT1XX & ",""FC""," & PartNumberColumn_Triple_B & "8:" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 & ")"
        Range(Cells(UR5 + 4, Column_Formato + 2), Cells(UR5 + 4, Column_Formato + 2)).Formula = "=" & PartNumberColumn_Triple_B & Ultimo_Renglon_CECE + 8 + 2
              
        'Copeo de formulas
        Renglon_T1XX = 8
        NumdeParte = Column_Ident - 12
        Columna_puntos_form = 11 + NumdeParte + 1 + (NumdeParte * 4) + 2
        PartNumberPuntos1100 = Application.WorksheetFunction.VLookup(Columna_puntos_form, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        PartNumberMat1100 = Application.WorksheetFunction.VLookup(Columna_puntos_form + 1, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        PartNumberCopper1100 = Application.WorksheetFunction.VLookup(Columna_puntos_form + 2, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        ColumnaFormatoParte = CECE_column - 5
        PartNumberOcurrencias = Application.WorksheetFunction.VLookup(ColumnaFormatoParte, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        
        
        For a = 1 To Ultimo_Renglon_CECE
            With Workbooks(CECE_file).Worksheets("T1XX_Format")
                Renglon_T1XX = a + 7
                Range(Cells(a + 7, Column_Formato), Cells(a + 7, Column_Formato)).Formula = "=" & PartNumberOcurrencias & Renglon_T1XX & "*(" & PartNumberPuntos1100 & Renglon_T1XX & "/ 100)"
                Range(Cells(a + 7, Column_Formato + 1), Cells(a + 7, Column_Formato + 1)).Formula = "=" & PartNumberOcurrencias & Renglon_T1XX & "*" & PartNumberMat1100 & Renglon_T1XX
                Range(Cells(a + 7, Column_Formato + 2), Cells(a + 7, Column_Formato + 2)).Formula = "=" & PartNumberOcurrencias & Renglon_T1XX & "*" & PartNumberCopper1100 & Renglon_T1XX
                
            End With
        Next a
                                
    CECE_column = CECE_column + 1
    Column_Formato = Column_Formato + 4
    'If CECE_column = 90 Then Stop
        
    Loop Until Workbooks(CECE_file).Sheets("CECE").Cells(24, CECE_column).Value = Empty
        'Ultima parte del machote
        PartNumberColumn = Application.WorksheetFunction.VLookup(Column_Formato + 1, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        ThisWorkbook.Worksheets("Machote").Range("S7:X7").Copy
        With Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & 7)
            .Cells(1).PasteSpecial xlPasteColumnWidths
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
            '.Cells(1).Select
            Application.CutCopyMode = False
        End With
  
        PartNumberColumn = Application.WorksheetFunction.VLookup(Column_Formato + 1, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        PartNumberColumn_B = Application.WorksheetFunction.VLookup(Column_Formato + 6, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & "8:" & PartNumberColumn_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeTop).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & "8:" & PartNumberColumn_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeBottom).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & "8:" & PartNumberColumn_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeLeft).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & "8:" & PartNumberColumn_B & Ultimo_Renglon_CECE + 8).Borders(xlEdgeRight).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & "8:" & PartNumberColumn_B & Ultimo_Renglon_CECE + 8).Borders(xlInsideHorizontal).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range(PartNumberColumn & "8:" & PartNumberColumn_B & Ultimo_Renglon_CECE + 8).Borders(xlInsideVertical).Weight = xlThin
                                                           
                                                           
        'Copia material
        part_number_column CECE_file, PartNumberColumn_2
        'ColumnaMaterial = PartNumberColumn_2 - 29
        'PartNumberColumn_Material = Application.WorksheetFunction.VLookup(ColumnaMaterial, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        'PartNumberColumn_Material_FormatT1XX = Application.WorksheetFunction.VLookup(Column_Formato + 2, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        'Workbooks(CECE_file).Sheets("CECE").Range(PartNumberColumn_Material & "29:" & PartNumberColumn_Material & Ultimo_Renglon_CECE + 28).Copy
        'With Workbooks(CECE_file).Sheets("T1XX_Format").Range(PartNumberColumn_Material_FormatT1XX & 8)
        '    .Cells(1).PasteSpecial xlPasteValues, , False, False
        '    .NumberFormat = "0.0000_) ;[red](0.0000)"
            '.Cells(1).PasteSpecial xlPasteFormats, , False, False
        'End With
          
        'Copia puntos
        ColumnaPoints = PartNumberColumn_2 - 35
        PartNumberColumn_Points = Application.WorksheetFunction.VLookup(ColumnaPoints, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        Workbooks(CECE_file).Sheets("CECE").Range(PartNumberColumn_Points & "29:" & PartNumberColumn_Points & Ultimo_Renglon_CECE + 28).Copy
        With Workbooks(CECE_file).Sheets("T1XX_Format").Range(PartNumberColumn & 8)
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
        End With
          
        'Copper
        ColumnaCopperCECE = PartNumberColumn_2 - 22
        PartNumberColumn_Column_Copper_CECE = Application.WorksheetFunction.VLookup(ColumnaCopperCECE, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        PartNumberColumn_Column_Copper_FormatT1XX = Application.WorksheetFunction.VLookup(Column_Formato + 3, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        Workbooks(CECE_file).Sheets("CECE").Range(PartNumberColumn_Column_Copper_CECE & "29:" & PartNumberColumn_Column_Copper_CECE & Ultimo_Renglon_CECE + 28).Copy
        With Workbooks(CECE_file).Sheets("T1XX_Format").Range(PartNumberColumn_Column_Copper_FormatT1XX & 8)
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
        End With
          
        'Material y Transporte
        ColumnaMat_Trans_CECE = PartNumberColumn_2 - 29
        PartNumberColumn_Column_Mat_Trans_CECE = Application.WorksheetFunction.VLookup(ColumnaMat_Trans_CECE, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        PartNumberColumn_Column_Mat_Trans_FormatT1XX = Application.WorksheetFunction.VLookup(Column_Formato + 6, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        Workbooks(CECE_file).Sheets("CECE").Range(PartNumberColumn_Column_Mat_Trans_CECE & "29:" & PartNumberColumn_Column_Mat_Trans_CECE & Ultimo_Renglon_CECE + 28).Copy
        With Workbooks(CECE_file).Sheets("T1XX_Format").Range(PartNumberColumn_Column_Mat_Trans_FormatT1XX & 8)
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .NumberFormat = "0.0000_) ;[red](0.0000)"
            '.Cells(1).PasteSpecial xlPasteFormats, , False, False
        End With
                
        'Agrega formula para Raw Material
        PartNumberColumn_Material_FormatT1XX = Application.WorksheetFunction.VLookup(Column_Formato + 6, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        For Renglon_B = 1 To Ultimo_Renglon_CECE
            With Workbooks(CECE_file).Worksheets("T1XX_Format")
                Range(Cells(Renglon_B + 7, Column_Formato + 2), Cells(Renglon_B + 7, Column_Formato + 2)).Formula = "=" & PartNumberColumn_Column_Mat_Trans_FormatT1XX & Renglon_B + 7 & "/1.05"
                Range(Cells(Renglon_B + 7, Column_Formato + 2), Cells(Renglon_B + 7, Column_Formato + 2)).NumberFormat = "0.0000_) ;[red](0.0000)"
                'Range(PartNumberColumn_Column_Mat_Trans_FormatT1XX & Renglon_B + 7).NumberFormat = "0.0000_) ;[red](0.0000)"
                valuetoredondear = Range(PartNumberColumn_Column_Mat_Trans_FormatT1XX & Renglon_B + 7).Value
                ValorRedondeado = Round(valuetoredondear, 4)
                Range(PartNumberColumn_Column_Mat_Trans_FormatT1XX & Renglon_B + 7) = ValorRedondeado
            End With
        Next Renglon_B
          
        'Pone cuadros en todo el machote estatico
        Ultimo_Renglon_CECE = Workbooks(CECE_file).Worksheets("CECE").Range("C29", Worksheets("CECE").Range("C29").End(xlDown)).Rows.Count
        PartNumberColumn = Application.WorksheetFunction.VLookup(Column_Ident, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A8:" & PartNumberColumn & Ultimo_Renglon_CECE + 8).Borders(xlEdgeTop).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A8:" & PartNumberColumn & Ultimo_Renglon_CECE + 8).Borders(xlEdgeBottom).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A8:" & PartNumberColumn & Ultimo_Renglon_CECE + 8).Borders(xlEdgeLeft).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A8:" & PartNumberColumn & Ultimo_Renglon_CECE + 8).Borders(xlEdgeRight).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A8:" & PartNumberColumn & Ultimo_Renglon_CECE + 8).Borders(xlInsideHorizontal).Weight = xlThin
        Workbooks(CECE_file).Worksheets("T1XX_Format").Range("A8:" & PartNumberColumn & Ultimo_Renglon_CECE + 8).Borders(xlInsideVertical).Weight = xlThin
        
    'Ciclo de llenado
    
    Renglon_CECE = 29
    Renglon_Formato_T1XX = 8
    Do
    'Parte 1 no dinamica
        With Workbooks(CECE_file)
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 2) = Sheets("CECE").Cells(Renglon_CECE, 2).Value
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 3) = Sheets("CECE").Cells(Renglon_CECE, 3).Value  ' DELPHI   P/N
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 5) = Sheets("CECE").Cells(Renglon_CECE, 5).Value  'DESCRIPTION
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 6) = Sheets("CECE").Cells(Renglon_CECE, 8).Value  'UOM mm, Pc, kg, Lt
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 7) = Sheets("CECE").Cells(Renglon_CECE, 9).Value  'Units
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 8) = Sheets("CECE").Cells(Renglon_CECE, 12).Value  '(A add), (D delete)
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 10) = Sheets("CECE").Cells(Renglon_CECE, 15).Value  'Type of component
            
            If Sheets("CECE").Cells(Renglon_CECE, 16).Value > 0 Then 'Cable
                Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 11) = Sheets("CECE").Cells(Renglon_CECE, 16).Value
            End If
            
            'Formatos centrado..etc
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 2).HorizontalAlignment = xlCenter
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 3).HorizontalAlignment = xlCenter
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 5).HorizontalAlignment = xlLeft
            Sheets("T1XX_Format").Columns(5).AutoFit
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 6).HorizontalAlignment = xlCenter
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 7).HorizontalAlignment = xlCenter
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 8).HorizontalAlignment = xlCenter
            Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, 10).HorizontalAlignment = xlCenter
            
        End With
    
    'Parte 2 Dinamica
        Columna_CECE_A = 17
        Columna_Formato_T1XX = 12
        With Workbooks(CECE_file)
            Do
                If Sheets("CECE").Cells(Renglon_CECE, Columna_CECE_A).Value > 0 Then 'Cable
                    Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, Columna_Formato_T1XX) = Sheets("CECE").Cells(Renglon_CECE, Columna_CECE_A).Value
                    Sheets("T1XX_Format").Cells(Renglon_Formato_T1XX, Columna_Formato_T1XX).HorizontalAlignment = xlCenter
                End If
                Columna_CECE_A = Columna_CECE_A + 1
                Columna_Formato_T1XX = Columna_Formato_T1XX + 1
            Loop Until Workbooks(CECE_file).Sheets("CECE").Cells(24, Columna_CECE_A).Value = Empty
        End With
                        
        Renglon_CECE = Renglon_CECE + 1
        Renglon_Formato_T1XX = Renglon_Formato_T1XX + 1
    Loop Until Workbooks(CECE_file).Sheets("CECE").Cells(Renglon_CECE, 3).Value = Empty
    
   Application.CutCopyMode = False
  


End Sub


Sub T1XX_Clasifier()
        'Solo pone NC a FC a T1XX
        
        LastcolumnCECEPartNumber = 12: rowPart = 6
        Do
            LastcolumnCECEPartNumber = LastcolumnCECEPartNumber + 1
        Loop While Not Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(rowPart, LastcolumnCECEPartNumber) = Empty
        LastcolumnCECEPartNumber = LastcolumnCECEPartNumber - 1
        
        Row_1 = 8

        Do
        
        'If Row_1 = 503 Then Stop
            Cable_Gauge_1 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 11).Value
            Tape_Cond = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 10).Value
            If Cable_Gauge_1 > 0 Or Tape_Cond = "TAPE" Or Tape_Cond = "COND" Then
                Row_2 = Row_1 + 1
                Item_1 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 2).Value
                Mspec_Value_1 = Left(Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 3).Value, 5)
                'If SumadeOcurrencias > 1 Then
                
                    Do
                        'AplicaOcurrencias_1 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 12)
                        'AplicaOcurrencias_2 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 12)
                        'sumadeocurrencias = AplicaOcurrencias_1 + AplicaOcurrencias_2
                        
                        'If sumadeocurrencias = 0 Then
                        '    AplicaOcurrencias_1 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 13)
                        '    AplicaOcurrencias_2 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 13)
                        '    sumadeocurrencias = AplicaOcurrencias_1 + AplicaOcurrencias_2
                        
                        'End If

                        Item_2 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 2).Value
                        Mspec_Value_2 = Left(Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 3).Value, 5)
                        Cable_Gauge_2 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 11).Value
                        'If Item_1 = Item_2 And Mspec_Value_1 = Mspec_Value_2 Then  'And Cable_Gauge_2 < 0
                        If Mspec_Value_1 = Mspec_Value_2 Then  'And Cable_Gauge_2 < 0
                                Units_Length_1 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 7).Value
                                Units_Length_2 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 7).Value
                                Units_Delta = Abs(Units_Length_1 - Units_Length_2)

                                Columnaq = 12
                                'If row_1 = 471 Then Stop
                                Do
                                    AplicaOcurrencias_1 = Empty: AplicaOcurrencias_2 = Empty: sumadeocurrencias = Empty
                                    AplicaOcurrencias_1 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, Columnaq)
                                    AplicaOcurrencias_2 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, Columnaq)
                                    sumadeocurrencias = AplicaOcurrencias_1 + AplicaOcurrencias_2
                                    Columnaq = Columnaq + 1
                                    'If Columnaq > 100 Then Stop
                                Loop While sumadeocurrencias < 2 And Columnaq < LastcolumnCECEPartNumber + 1 And Columnaq < 150


                                If sumadeocurrencias > 1 Then
                                    Mspec_1 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 3).Value
                                    Mspec_2 = Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 3).Value
                                    If Mspec_1 = Mspec_2 Then
                                        True_False = "True"
                                    Else
                                        True_False = "False"
                                    End If
                                    
                                    Dim String1 As String           'Este bloque se debe integrar al proceso principal se agrego asi para hacer simple su implementacion
                                    Dim String2 As String
                                    Dim Arr1 As Variant
                                    Dim Arr2 As Variant
                                    On Error Resume Next
                                    String1 = Item_1
                                    String2 = Item_2
                                    Arr1 = Split(String1, ":")
                                    Arri1_mod = Arr1(0)
                                    Arr2 = Split(String2, ":")
                                    Arri2_mod = Arr2(0)

                                    
                                    Select Case Tape_Cond
                                        Case Is = "COND"
                                            If Units_Delta < 200 And True_False = "True" And Arri1_mod = Arri2_mod And sumadeocurrencias > 1 Then
                                                Borrar = "Si"
                                            End If
                                        
                                        Case Is = "TAPE"
                                            If Units_Delta < 200 And True_False = "True" And Arri1_mod = Arri2_mod And sumadeocurrencias > 1 Then
                                                Borrar = "Si"
                                            End If
                                            String1 = Empty: String2 = Empty: Arr1 = Empty: Arr2 = Empty: Arri1_mod = Empty: Arri2_mod = Empty
                                        
                                        Case Is = "CABL"
                                            Escala_Cable_1 = Application.WorksheetFunction.VLookup(Cable_Gauge_1, ThisWorkbook.Worksheets("T1XX Cables").Range("C4:D14"), 2)
                                            Escala_Cable_2 = Application.WorksheetFunction.VLookup(Cable_Gauge_2, ThisWorkbook.Worksheets("T1XX Cables").Range("C4:D14"), 2)
                                            Delta_Escala = Abs(Abs(Escala_Cable_1) - Abs(Escala_Cable_2))
                                                                    
                                            If Units_Delta < 200 And Delta_Escala < 2 And Arri1_mod = Arri2_mod And sumadeocurrencias > 1 Then
                                                Borrar = "Si"
                                            End If
                                            
                                    End Select
                                    String1 = Empty: String2 = Empty: Arr1 = Empty: Arr2 = Empty: Arri1_mod = Empty: Arri2_mod = Empty
                                End If
                        End If
 
                                                                        
                        
                        If Not Borrar = "Si" Then
                            Item_2 = Empty
                            Mspec_Value_2 = Empty
                            Cable_Gauge_2 = Empty
                            True_False = Empty
                            Mspec_1 = Empty
                            Mspec_2 = Empty
                            Row_2 = Row_2 + 1
                        End If
                        AplicaOcurrencias_1 = Empty
                        AplicaOcurrencias_2 = Empty
                        sumadeocurrencias = Empty
                    Loop While Not Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 3).Value = Empty And Not Borrar = "Si"

                'End If
                
                Else
                'Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1) = "FC"
            End If

            'With Workbooks(CECE_file).Worksheets("T1XX_Format")
            '    If Cells(Row_1, 8) = "D" Then
            '        Cells(Row_1, Column_Formato + 2) = Cells(Row_1, Column_Formato + 2) * -1
            '    End If
            'End With
            
            'If Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 3).Value = "TAPESTART" Or Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 3).Value = "STARTTAPE" Then
                'Workbooks(CECE_file).Worksheets("T1XX_Format").Rows(Row_1).EntireRow.Delete
                'Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1) = "NC"
                'Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1).HorizontalAlignment = xlLeft
                'Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1).Font.Bold = True
                'Row_1 = Row_1 + 1
            'Else
            
                If Borrar = "Si" Then
                    'Workbooks(CECE_file).Worksheets("CECE").Rows(Row_2).EntireRow.Delete
                    'Workbooks(CECE_file).Worksheets("CECE").Rows(Row_1).EntireRow.Delete
                    Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1) = "NC"
                    Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 1) = "NC"
                    Workbooks(CECE_file).Worksheets("CECE").Cells(Row_1 + 21, 1) = "NC"
                    Workbooks(CECE_file).Worksheets("CECE").Cells(Row_2 + 21, 1) = "NC"

                    
                    'If Row_2 = 509 Then Stop
                    Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1).HorizontalAlignment = xlLeft
                    Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1).Font.Bold = True
                    Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 1).HorizontalAlignment = xlLeft
                    Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_2, 1).Font.Bold = True
                                                            
                    
                    Row_1 = Row_1 + 1
                    
                Else
                    If Not Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1) = "NC" Then
                        Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1) = "FC"
                        Workbooks(CECE_file).Worksheets("CECE").Cells(Row_1 + 21, 1) = "FC"
                        Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1).HorizontalAlignment = xlLeft
                        Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 1).Font.Bold = True
                    End If
                    Row_1 = Row_1 + 1
                End If
            
            'End If
                        
            Borrar = Empty
            Item_1 = Empty
            Item_2 = Empty
            Units_Length_1 = Empty
            Units_Length_2 = Empty
            Cable_Gauge_1 = Empty
            Cable_Gauge_2 = Empty
            Tape_Cond = Empty
            'Poner negativos en puntos

        Loop While Not Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_1, 3).Value = Empty
End Sub


Sub Copiador_De_Cobre()

    Ultimo_Renglon_T1XX_Format = Workbooks(CECE_file).Worksheets("T1XX_Format").Range("C8", Worksheets("T1XX_Format").Range("C8").End(xlDown)).Rows.Count
    Columna_Price_Detail = 15
    Column_IdentA = Column_Ident + 3
    Nombre_Tab = Name_Tab       'Added on 4/11/17
    Do
        With Workbooks(CECE_file)
            'Sheets(Tipo).Cells(19, Columna_Price_Detail) = Sheets("T1XX_Format").Cells(Ultimo_Renglon_T1XX_Format + 7 + 7, Column_IdentA).Value
            a = Sheets("T1XX_Format").Cells(Ultimo_Renglon_T1XX_Format + 7 + 7, Column_IdentA).Value
            Sheets(Nombre_Tab).Cells(19, Columna_Price_Detail) = Sheets("T1XX_Format").Cells(Ultimo_Renglon_T1XX_Format + 7 + 7, Column_IdentA).Value
        End With
        Column_IdentA = Column_IdentA + 4
        Columna_Price_Detail = Columna_Price_Detail + 2
    Loop While Not Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Ultimo_Renglon_T1XX_Format + 7 + 7, Column_IdentA).Value = Empty

End Sub

Sub Deleter_T1XX_Tab()

    Row_T1XX = 8
    Counter_NC = Empty
    Counter_FC = Empty
    Do
        
        
        If Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_T1XX, 1).Value = "NC" Then
            Counter_NC = Counter_NC + 1
            
        End If
        
        If Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_T1XX, 1).Value = "FC" Then
            Counter_FC = Counter_FC + 1
        End If
        
        Row_T1XX = Row_T1XX + 1
    
    Loop While Not Workbooks(CECE_file).Worksheets("T1XX_Format").Cells(Row_T1XX, 1).Value = Empty
    
    'BatchProcesing = "Yes" es para cuando corre el batch y no se requieren alarmas
    If Counter_NC + 8 = Row_T1XX Or Counter_FC + 8 = Row_T1XX Then
        If Counter_NC + 8 = Row_T1XX Then
            Workbooks(CECE_file).Worksheets("CECE").Cells(22, 2) = "NC"
        Else
            Workbooks(CECE_file).Worksheets("CECE").Cells(22, 2) = "FC"
        End If
        
        Workbooks(CECE_file).Worksheets("T1XX_Format").Delete
        Bandera_Activar_Cobre = "No"
        
        Checkbox9 = ThisWorkbook.Worksheets("Sheet1").Shapes("Check Box 9").ControlFormat.Value
            If Checkbox9 = 1 Then
                If BatchProcesing = Empty Then
                    Warning327.Show
                End If
        End If
    End If
    


End Sub
'Cahnges to line 184 DisplayAlerts Added on 4/10/17
'Changes to lines 12/164/555/558//578/788/1505 Name_Tab
'Line 734 updated the error code 4/12/17
'Line 181 Module3.Redondear 'disabled per Roger York 4/13/17

Sub calculador_De_Cobre()

Renglon = 29
NuevoCobre = ThisWorkbook.Worksheets("Sheet1").Cells(21, 3).Value
CECE_file = ActiveWorkbook.Name
part_number_column CECE_file, PartNumberColumn_2
CopperColumn = PartNumberColumn_2 + 14

Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.ScreenUpdating = False



CopperColumnLetter = Application.WorksheetFunction.VLookup(CopperColumn + 2, ThisWorkbook.Worksheets("Sheet4").Range("A9:B414"), 2)
CopperLBSColumnLetter = Application.WorksheetFunction.VLookup(CopperColumn - 1, ThisWorkbook.Worksheets("Sheet4").Range("A9:B414"), 2)

Do
    CopperActual = Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, CopperColumn - 1).Value
    If CopperActual > 0 Then
    
        SaveOriginalUnitCost = Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, CopperColumn).Value
        TPCostWithNewCopperRate = SaveOriginalUnitCost + (CopperActual * NuevoCobre)
        'Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, CopperColumn + 2) = Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, CopperColumn).Value
        'Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, CopperColumn + 1).Formula = "=+DZ29+(DU29*2.5)"
        'Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, CopperColumn + 1).Formula = "=+" & CopperColumnLetter & Renglon & "+(" & CopperLBSColumnLetter & Renglon & "*2.5)"
        Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 1) = TPCostWithNewCopperRate
        Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 1).Interior.ColorIndex = 42
        Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 4) = TPCostWithNewCopperRate
        Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 4).Interior.ColorIndex = 42
        Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 5) = TPCostWithNewCopperRate
        Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 5).Interior.ColorIndex = 42
        'Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 6) = Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, CopperColumn + 1).Value
        'Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, PartNumberColumn_2 + 6).Interior.ColorIndex = 42
        
        CopperActual = Empty
    End If
    Renglon = Renglon + 1
Loop Until Workbooks(CECE_file).Sheets("CECE").Cells(Renglon, 3).Value = Empty
Workbooks(CECE_file).Sheets("CECE").Range("L3") = NuevoCobre
Workbooks(CECE_file).Sheets("CECE").Range("L2") = "Copper updated by sales"
Application.Calculation = xlAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub

Sub corrector_de_costos()

CECE_file = ActiveWorkbook.Name
Checador_de_EWO CECE_file, EndFlag
If EndFlag = "Terminar" Then Exit Sub
PrecioDeCobre = ThisWorkbook.Worksheets("Sheet1").Cells(21, 3).Value

part_number_column CECE_file, PartNumberColumn_2


Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.ScreenUpdating = False

RenglonCECE = 29
Do
    CECE_PartNumber = Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, 3)
    CECE_PartNumberMod = Val(CECE_PartNumber)
    
    '-----Convertidor de texto a numero cuando sea nesesario para el vlookup
    If CECE_PartNumberMod = 0 Then
    Else
        CECE_PartNumber = CECE_PartNumberMod
    End If
    
    On Error Resume Next
    ExiteLista = Application.WorksheetFunction.VLookup(CECE_PartNumber, ThisWorkbook.Worksheets("Sheet5").Range("B2:C100"), 1, False)
    If Not ExiteLista = Empty Then
        CostoSinCobre = Application.WorksheetFunction.VLookup(CECE_PartNumber, ThisWorkbook.Worksheets("Sheet5").Range("B2:C100"), 2, False)
            CopperWeight = Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 13).Value
            CostoConCobre = CostoSinCobre + (CopperWeight * PrecioDeCobre)
            
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 1) = CostoConCobre
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 1).Interior.ColorIndex = 42
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 4) = CostoConCobre
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 4).Interior.ColorIndex = 42
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 5) = CostoConCobre
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 5).Interior.ColorIndex = 42
            'Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 14) = NewValue
            'Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 14).Interior.ColorIndex = 42
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 10) = "Cost updated by sales"
            Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 10).Interior.ColorIndex = 42
            CECE_PartNumber = Empty
            ExiteLista = Empty
            NewValue = Empty
    End If
    RenglonCECE = RenglonCECE + 1
Loop Until Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, 2).Value = Empty

Application.Calculation = xlAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True

FinishedCECE.Show

End Sub

Sub Checador_de_EWO(CECE_file, EndFlag)
    
        On Error Resume Next
        Primer_Ident = Workbooks(CECE_file).Worksheets("CECE").Cells(8, 2).Value
        Segundo_Ident = Workbooks(CECE_file).Worksheets("CECE").Cells(20, 2).Value
        If Primer_Ident = "Cost Estimating Change Evaluation" And Segundo_Ident = "Purchasing Customer:" Then
            EndFlag = Empty
        Else
            EndFlag = "Terminar"
            WarningActiveCECE.Show
        End If
    
End Sub

Sub ReporteDeEWOCostosFueraDeRango()

    ExcelOutput = ActiveWorkbook.Name
    Hoja = ActiveSheet.Name
    ReportRow = 3
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    
    'Clean Page
    Do
        ComponentPN = ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 2).Value
        If Not ComponentPN = Empty Then
            'ThisWorkbook.Worksheets("Sheet5").Range(Cells(ReportRow, 2), Cells(ReportRow, 8)).Select
            
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 2).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 3).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 3).Interior.ColorIndex = xlNone
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 4).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 5).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 6).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 7).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 8).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 9).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 10).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 11).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 12).ClearContents
            ThisWorkbook.Worksheets("Sheet5").Cells(ReportRow, 13).ClearContents
            Selection.ClearContents
           
            
        End If
        ReportRow = ReportRow + 1
        ComponentPN = Empty
    Loop Until ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 2).Value = Empty
    
       
    
    NuevoCobre = ThisWorkbook.Worksheets("Sheet1").Cells(21, 3).Value
    ExcelOutputRow = 4
    ReportRow = 3
    Do
        CR = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 5).Value
        ComponentPN = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 11).Value
        ComponentPrice = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 17).Value
        CopperPrice = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 16).Value
        Cost = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 17).Value
        MinCost = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 18).Value
        MaxCost = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 19).Value
        Descripcion = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 12).Value
        CopperWeight = Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 15).Value
        
        '13752406
        'If ComponentPN = 13752406 Then Stop
        
        
        If Cost < MinCost Or Cost > MaxCost Then
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 8) = CR
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 2) = ComponentPN
            'Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, 8) = CopperPrice
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 7) = Cost
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 5) = MinCost
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 6) = MaxCost
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 3) = ComponentPrice
            '-----Pone delta
            
        Select Case Cost
            Case Is > MaxCost
                ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 9) = Round(Cost - MaxCost, 4)
            Case Is < MinCost
                ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 9) = Round(MinCost - Cost, 4)
        End Select
            
           
            
            media = Round((MinCost + MaxCost) / 2, 4)
            TercerQuarto = (MaxCost + media) / 2
            'calculo de costo final, menos el cobre
            CostoenFormato = TercerQuarto - (CopperWeight * NuevoCobre)
            If CostoenFormato < 0 Then
                ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 4).Interior.ColorIndex = 3
                ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 4) = Cost
                ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 13) = "Costo negativo si se usa CCH"
            Else
                ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 4) = media
            End If
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 12) = Descripcion
            ReportRow = ReportRow + 1
        End If
        
        If Not NuevoCobre = CopperPrice Then
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 8) = CR
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 2) = ComponentPN
            ThisWorkbook.Sheets("Sheet5").Cells(ReportRow, 14) = CopperPrice
            ReportRow = ReportRow + 1
        End If
                    
        CR = Empty: ComponentPN = Empty: CopperPrice = Empty: Cost = Empty: MinCost = Empty: MaxCost = Empty: CostoenFormato = Empty: TercerQuarto = Empty
    
    ExcelOutputRow = ExcelOutputRow + 1
    Loop Until Workbooks(ExcelOutput).Worksheets(Hoja).Cells(ExcelOutputRow, 11).Value = Empty
    
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    
    
    EvalualTerminado.Show

End Sub


Sub Advertencia_costo_Cobre_diferentes_rapida(EndFlag)
    
    Dim sht As Worksheet
    Dim LastRow As Long
    Set sht = Workbooks(CECE_file).Worksheets("CECE")
    
    StartTime3 = Timer
    part_number_column CECE_file, PartNumberColumn_2
    SecondsElapsed3 = Round(Timer - StartTime3, 2)

    RenglonCECE = 29
    
    LastRow = sht.Range("C29").CurrentRegion.Rows.Count
    CECE_PartNumber = Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, 3)
    
    On Error Resume Next
    'ExiteLista = Application.WorksheetFunction.VLookup(CECE_PartNumber, ThisWorkbook.Worksheets("Sheet5").Range("C23:C" & LastRow), 1, False)
    'RenglonEncontrado = Application.WorksheetFunction.Match(CECE_PartNumber, ThisWorkbook.Worksheets("Sheet5").Range("C2" & RenglonCECE & ":C" & LastRow), 0)
    
    Do
        CECE_PartNumberPivot = Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, 3)
        renglonpivot_B = Application.WorksheetFunction.Match(CECE_PartNumberPivot, Workbooks(CECE_file).Sheets("CECE").Range("C" & RenglonCECE + 1 & ":C" & LastRow), 0)
        
        If Not renglonpivot_B = Empty Then
            renglonpivot_B = renglonpivot_B + RenglonCECE
            Cost_Pivot_A = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 4).Value)
            Cost_Pivot_B = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(renglonpivot_B, PartNumberColumn_2 + 4).Value)
                
            Copper_Pivot_A = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 13).Value)
            Copper_Pivot_B = Abs(Workbooks(CECE_file).Worksheets("CECE").Cells(renglonpivot_B, PartNumberColumn_2 + 13).Value)
                
            GAUGE_Pivot_A = Workbooks(CECE_file).Worksheets("CECE").Cells(RenglonCECE, 16).Value
            
            Delta_Cost = Cost_Pivot_A - Cost_Pivot_B
            Delta_Copper = Copper_Pivot_A - Copper_Pivot_B
            
                'Cost
            If Delta_Cost <> 0 And GAUGE_Pivot_A = Empty Then
                
                With Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, 3).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    
                With Workbooks(CECE_file).Sheets("CECE").Cells(renglonpivot_B, 3).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                
                With Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 4).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    
                With Workbooks(CECE_file).Sheets("CECE").Cells(renglonpivot_B, PartNumberColumn_2 + 4).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    
                EndFlag = "Terminar"
                EndFlag_2 = "Terminar ciclo"
            End If
                    
            'Copper
            If Delta_Copper <> 0 And GAUGE_Pivot_A = Empty Then
                
                With Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, 3).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    
                With Workbooks(CECE_file).Sheets("CECE").Cells(renglonpivot_B, 3).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                                                
                With Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE, PartNumberColumn_2 + 13).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                                
                With Workbooks(CECE_file).Sheets("CECE").Cells(renglonpivot_B, PartNumberColumn_2 + 13).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                EndFlag = "Terminar"
                EndFlag_2 = "Terminar ciclo"
            End If
        End If
            
        renglonpivot_B = Empty
        Numero_Pivot_A = Empty
        Numero_Pivot_B = Empty
        Cost_Pivot_A = Empty
        Cost_Pivot_B = Empty
        GAUGE_Pivot_A = Empty
  
        RenglonCECE = RenglonCECE + 1
    Loop Until Workbooks(CECE_file).Sheets("CECE").Cells(RenglonCECE + 1, 2).Value = Empty
        
    If EndFlag = "Terminar" Then
        AdventenciaCosto.Show
    End If

End Sub

Sub CopiadorMejor()

Dim rFind As Range


NombredeTab = "CR " & Cr_number & " 200mm & Gage Rule"
Name_Tab = "CR" & Cr_number & "-GM Detail"

Columna = 14

With Workbooks(CECE_file).Sheets(NombredeTab).Range("A1:ALK7000")
    Set rFind = .Find(what:="IN FINAL QUOTE (FC)", LookAt:=xlWhole, MatchCase:=False, searchformat:=False)
    If Not rFind Is Nothing Then
        a = rFind.Row
        b = rFind.Column
    End If
End With

Last_LetterColumn = Application.WorksheetFunction.VLookup(b, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
Do
    GMDetailPart = Empty: T1XXCoumnpart = Empty: CopperTotal = Empty: RawMaterial = Empty: LaborTotal = Empty
    GMDetailPart = Workbooks(CECE_file).Sheets(Name_Tab).Cells(21, Columna).Value
    'T1XXCoumnpart = Application.WorksheetFunction.Match(GMDetailPart, Workbooks(CECE_file).Worksheets(NombredeTab).Range("R6:II6"), 0) + 17 + 2
    On Error Resume Next
    
    T1XXCoumnpart = Application.WorksheetFunction.Match(GMDetailPart, Workbooks(CECE_file).Worksheets(NombredeTab).Range(Last_LetterColumn & "6:ALK6"), 0)
    T1XXCoumnpart = b - 1 + T1XXCoumnpart
    LaborTotal = Workbooks(CECE_file).Sheets(NombredeTab).Cells(a, T1XXCoumnpart).Value
    RawMaterial = Workbooks(CECE_file).Sheets(NombredeTab).Cells(a, T1XXCoumnpart + 1).Value
    CopperTotal = Workbooks(CECE_file).Sheets(NombredeTab).Cells(a, T1XXCoumnpart + 2).Value
    
    
    GMDetail_LetterColumn = Application.WorksheetFunction.VLookup(Columna + 1, ThisWorkbook.Worksheets("Sheet4").Range("A9:B1000"), 2)
    'Range("O6:O19").Select
    'Desabilitado por platica con July Wong 2/28/18
    'a = GMDetail_LetterColumn & "6:" & GMDetail_LetterColumn & "19"
    'Workbooks(CECE_file).Worksheets(Name_Tab).Range(GMDetail_LetterColumn & "6:" & GMDetail_LetterColumn & "19").Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Application.CutCopyMode = False
    
    'CopperTotal = Round(CopperTotal, 4)
    
    Workbooks(CECE_file).Sheets(Name_Tab).Cells(7, Columna + 1) = RawMaterial
    Workbooks(CECE_file).Sheets(Name_Tab).Cells(8, Columna + 1) = LaborTotal
    Workbooks(CECE_file).Sheets(Name_Tab).Cells(19, Columna + 1) = CopperTotal

    '"IN FINAL QUOTE (FC)"

    Columna = Columna + 2
Loop Until Workbooks(CECE_file).Sheets(Name_Tab).Cells(21, Columna).Value = Empty


End Sub

Sub CopiarHarness() '6-25-18

    
        Columna_1 = 17
    'ThisWorkbook.Worksheets("Temp-1").Range("B2:B4140").Clear
        Do
            If Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna_1).Value = Empty Then
                Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna_1).Value = Workbooks(CECE_file).Worksheets("CECE").Cells(24, Columna_1).Text
            End If
            
            'Agregado 9/21/18 x caso de Ramon Dominguez cr 105539838
            If Workbooks(CECE_file).Worksheets("CECE").Cells(20, Columna_1).Value = Empty Then
                Workbooks(CECE_file).Worksheets("CECE").Cells(20, Columna_1).Value = Workbooks(CECE_file).Worksheets("CECE").Cells(22, Columna_1).Text
            End If
            
            
        Columna_1 = Columna_1 + 1
        Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(24, Columna_1).Value = Empty
        
    

End Sub


Sub borraTapeStart()

        Dim Hoja As Range
        Dim Col8 As Integer
        Dim Conteo As Integer
        Dim ColumnLetter As String
        Dim ContFilt As Integer
        Col8 = 100
        Do
            Detector = Empty: Col8 = Col8 + 1
            Detector = Workbooks(CECE_file).Worksheets("CECE").Cells(28, Col8).Text
            'Col8 = Col8 + 1
        Loop Until Detector = "Unit Cost w/o Cu" Or Col8 > 500
        Col8 = Col8 + 3
        ColumnLetter = Split(Cells(1, Col8).Address, "$")(1)
        
        Conteo = WorksheetFunction.CountA(Range("C28", Range("C28").End(xlDown))) + 28
        
        
        Application.DisplayAlerts = False
        
        Set Hoja = Workbooks(CECE_file).Worksheets("CECE").Range("B28:" & ColumnLetter & Conteo)
        Application.ScreenUpdating = True

'        With Info
'            .AutoFilter field:=3, Criteria1:="TAPESTART"
'        .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Rows.Delete
'        End With
        
        Hoja.Sort key1:=Range("C28"), order1:=xlAscending, Header:=xlYes
        Hoja.AutoFilter field:=2, Criteria1:="TAPESTART"
        
        ContFilt = Hoja.SpecialCells(xlCellTypeVisible).Rows.Count
        
        If ContFilt > 1 Then
            Hoja.Offset(1, 0).Resize(Hoja.Count - 1).SpecialCells(xlCellTypeVisible).Rows.Delete
        End If
        
        'Workbooks(CECE_file).Worksheets("CECE").AutoFilter = False
        Workbooks(CECE_file).Worksheets("CECE").ShowAllData
                        
        
        
End Sub



Sub GeneradorDeSubfijos()

    Application.StatusBar = True
    DoEvents
    Application.ScreenUpdating = True
    Application.StatusBar = "Adding Subfixes for proliferation"
    Application.ScreenUpdating = False
    DoEvents
    
    Conteo = Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columns.Count).End(xlToLeft).Column
    Columna = 17
    Do
        numero = Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna).Value
        Letra = ThisWorkbook.Worksheets("Sheet6").Cells(Columna - 7, 3).Text
        Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna) = numero & Letra
        numero2 = Workbooks(CECE_file).Worksheets("CECE").Cells(24, Columna).Value
        Workbooks(CECE_file).Worksheets("CECE").Cells(24, Columna) = numero2 & Letra
        numero3 = Workbooks(CECE_file).Worksheets("CECE").Cells(22, Columna).Value
        Workbooks(CECE_file).Worksheets("CECE").Cells(22, Columna) = numero3 & Letra
        numero4 = Workbooks(CECE_file).Worksheets("CECE").Cells(20, Columna).Value
        Workbooks(CECE_file).Worksheets("CECE").Cells(20, Columna) = numero4 & Letra
                
        Columna = Columna + 1
    Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna) = Empty
    


End Sub


Sub BorradorDeSubfijos()

    
    DoEvents
    Application.ScreenUpdating = True
    Application.StatusBar = True
    Application.StatusBar = "Deleting Subfixes used for proliferation"
    Application.ScreenUpdating = False
    DoEvents

    'conteo = Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columns.Count).End(xlToLeft).Column
    'Columna = 17
    
    'Do  'borra los subfijos en el CECE
     '   numero1 = Empty: numero2 = Empty: numero3 = Empty: numero4 = Empty
        
     '   numero1 = Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna).Value
     '   If Columna > 17 + 26 Then
     '       numero1 = Left(numero1, Len(numero1) - 2)
     '   Else
     '       numero1 = Left(numero1, Len(numero1) - 1)
     '   End If
     '   Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna) = numero1
        
     '   numero2 = Workbooks(CECE_file).Worksheets("CECE").Cells(24, Columna).Value
     '   If Columna > 17 + 26 Then
     '       numero2 = Left(numero2, Len(numero2) - 2)
     '   Else
     '       numero2 = Left(numero2, Len(numero2) - 1)
     '   End If
     '   Workbooks(CECE_file).Worksheets("CECE").Cells(24, Columna) = numero2
        
     '   numero3 = Workbooks(CECE_file).Worksheets("CECE").Cells(22, Columna).Value
     '   If Columna > 17 + 26 Then
     '       numero3 = Left(numero3, Len(numero3) - 2)
    '    Else
    '        numero3 = Left(numero3, Len(numero3) - 1)
     '   End If
    '    Workbooks(CECE_file).Worksheets("CECE").Cells(22, Columna) = numero3 & Letra
        
    '    numero4 = Workbooks(CECE_file).Worksheets("CECE").Cells(20, Columna).Value
    '    If Columna > 17 + 26 Then
    '        numero4 = Left(numero4, Len(numero4) - 2)
    '    Else
    '        numero4 = Left(numero4, Len(numero4) - 1)
    '    End If
    '    Workbooks(CECE_file).Worksheets("CECE").Cells(20, Columna) = numero4

    
    '    Columna = Columna + 1
    'Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(26, Columna) = Empty

    Columna = 14
    Do
        NumFrom = Empty: NumTo = Empty
        
        NumFrom = Workbooks(CECE_file).Worksheets(Name_Tab).Cells(21, Columna).Value
        If Columna > 14 + 26 + 24 Then
            NumFrom = Left(NumFrom, Len(NumFrom) - 2)
        Else
            NumFrom = Left(NumFrom, Len(NumFrom) - 1)
        End If
        Workbooks(CECE_file).Worksheets(Name_Tab).Cells(21, Columna) = NumFrom
        
        NumTo = Workbooks(CECE_file).Worksheets(Name_Tab).Cells(22, Columna).Value
        If Columna > 14 + 26 + 24 Then
            NumTo = Left(NumTo, Len(NumTo) - 2)
        Else
            NumTo = Left(NumTo, Len(NumTo) - 1)
        End If
        Workbooks(CECE_file).Worksheets(Name_Tab).Cells(22, Columna) = NumTo
        
        
        
        Columna = Columna + 2
    Loop Until Workbooks(CECE_file).Worksheets(Name_Tab).Cells(22, Columna) = Empty
    Application.StatusBar = False
    
End Sub

Sub BorradorDeRenglonesVacios()

    'CECE_file
    Renglon8 = 28: iCount = 28
    Dim r As Range
    Col8 = 100
    Do
        Detector = Empty: Col8 = Col8 + 1
        Detector = Workbooks(CECE_file).Worksheets("CECE").Cells(28, Col8).Text
        'Col8 = Col8 + 1
    Loop Until Detector = "Unit Cost w/o Cu" Or Col8 > 500
    ColumnLetter = Split(Cells(1, Col8).Address, "$")(1)
        
    'Tanque = 5
    Do
        Set r = Workbooks(CECE_file).Worksheets("CECE").Range(ColumnLetter & Renglon8)
        Bandera8 = Empty
        If r.HasFormula = True Then
            iCount = iCount + 1
            Bandera8 = "Tiene Formula"
        End If
        
        
        a = Workbooks(CECE_file).Worksheets("CECE").Range("JN8000").HasFormula
        
        
        
        If Not Workbooks(CECE_file).Worksheets("CECE").Cells(Renglon8, 3) = Empty Then
            ItemCounter = ItemCounter + 1
        End If
        
        
        Renglon8 = Renglon8 + 1: Component = Empty
    Loop Until Renglon8 > 10000 'Or Bandera8 = Empty
    
    NewRow = ItemCounter + 79
    If iCount > 9000 Then
        'Workbooks(CECE_file).Worksheets("CECE").Rows(NewRow & ":1048576").Delete
        'Workbooks(CECE_file).Worksheets("CECE").Range("A1:G37").Clear
        'Workbooks(CECE_file).Worksheets("CECE").Range(ColumnLetter & NewRow & ":" & ColumnLetter & "1048576").Select
        'Selection.Clear
        'Selection.Delete
        'Rows("10:17").Select
        'Rows(NewRow & ":1048576").Select
        'Selection.Delete Shift:=xlUp
        
        'Rows("81:81").Select
'        Rows(NewRow & ":" & NewRow).Select
'        Range(Selection, Selection.End(xlDown)).Select
'        Selection.Delete Shift:=xlUp
        
        Range(ColumnLetter & NewRow).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Delete Shift:=xlUp
        
        
        
    End If

End Sub

Sub NoCostCr()
    test1 = Empty: test2 = Empty
    ColPN = WorksheetFunction.Match("Part Number", Workbooks(CECE_file).Worksheets("CECE").Range("A28:ZZ28"), 0)
    test1 = Workbooks(CECE_file).Worksheets("CECE").Cells(30, 2)
    test2 = Workbooks(CECE_file).Worksheets("CECE").Cells(29, ColPN)
    
    
    If test1 = Empty And test2 = Empty Then
        'part_number_column CECE_file, PartNumberColumn_2
        Workbooks(CECE_file).Worksheets("CECE").Cells(29, ColPN) = 1
        Workbooks(CECE_file).Worksheets("CECE").Cells(29, 2) = 1
        Workbooks(CECE_file).Worksheets("CECE").Cells(29, 3) = 1
        Workbooks(CECE_file).Worksheets("CECE").Cells(29, 5) = "BOM IS NOT AFFECTED"
        Workbooks(CECE_file).Worksheets("CECE").Cells(29, 8) = "PCS"
        Workbooks(CECE_file).Worksheets("CECE").Cells(29, 12) = "A"
        Col = 17
        Do
            Workbooks(CECE_file).Worksheets("CECE").Cells(29, Col) = 1
            Col = Col + 1
        Loop Until Workbooks(CECE_file).Worksheets("CECE").Cells(26, Col) = Empty
    End If
           
    
End Sub

Sub Inv_Amort(AmortPriceRow, AmortPriceCol)
    Col25 = 15
    Dim WB25 As Workbook
    Dim WS25 As Worksheet
    Dim WSCount As Byte
    Dim Hoja25 As Byte
    'Dim ID25 As String
    Dim Amort As Double
    
    Set WB25 = Workbooks(CECE_file)
    'Set WS25 = WB25.Worksheets(Tipo)
    On Error Resume Next
    WSCount = WB25.Worksheets.Count
    
    For Hoja25 = 1 To WSCount
        'ID25 = Empty
        ID25 = Workbooks(CECE_file).Worksheets(Hoja25).Cells(AmortPriceRow, AmortPriceCol).Value
        If ID25 = "Amortization Price" Then
            Set WS25 = WB25.Worksheets(Hoja25)
            Amort = WS25.Cells(AmortPriceRow, AmortPriceCol + 1).Value
            Do
                WB25.Worksheets(Tipo).Cells(11, Col25).Formula = WB25.Worksheets(Tipo).Cells(11, Col25).Formula & "+" & Amort
                Col25 = Col25 + 2
            Loop Until WB25.Worksheets(Tipo).Cells(23, Col25) = Empty
            
            Exit For
        End If
    Next
    
'    If ID25 = Empty Then
'        MsgBox ("Warning!!, no Amortization information found in CECE file, program will stop")
'        End
'    End If
   
    


End Sub

Sub TestAmort(Flag26, Hoja26, AmortPriceRow, AmortPriceCol)
    Col25 = 15
    Dim WB26 As Workbook
    Dim WS26 As Worksheet
    Dim WSCount As Byte
    'Dim Hoja26 As Byte
    'Dim ID26 As String
    Dim Amort As Double
    Dim RangoB As Range
    
    Set WB26 = Workbooks(CECE_file)
    
    'Set WS26 = WB26.Worksheets(Tipo)
    On Error Resume Next
    WSCount = WB26.Worksheets.Count
    
    For Hoja26 = 1 To WSCount
        With WB26.Worksheets(Hoja26).Range("A1:J50")
            Set rFind_1 = .Find(what:="Project Expenses", LookAt:=xlWhole, MatchCase:=False, searchformat:=False)
            If Not rFind_1 Is Nothing Then
                a = rFind_1.Row
                b = rFind_1.Column
                
                Set rFind_2 = .Find(what:="Amortization Price", LookAt:=xlWhole, MatchCase:=False, searchformat:=False)
                If Not rFind_2 Is Nothing Then
                    AmortPriceRow = rFind_2.Row
                    AmortPriceCol = rFind_2.Column
                    'Amort = WS26.Worksheets(Hoja26).Cells(AmortPriceRow, AmortPriceCol + 1).value
                    If IsNumeric(WB26.Worksheets(Hoja26).Cells(AmortPriceRow, AmortPriceCol + 1)) = True Then
                        Flag26 = "Ok"
                        Exit For
                    Else
                        Dim l As Long, t As Long
                        WB26.Worksheets(Hoja26).Activate
                        l = Cells(AmortPriceRow + 1, AmortPriceCol + 2).Left
                        t = Cells(AmortPriceRow + 1, AmortPriceCol + 2).Top
                        'ActiveSheet.Shapes.AddConnector(msoConnectorStraight, t + 89.25, l + 89.25, l, t).Select
                        ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 500, t, l, t).Select
                        With Selection
                            With .ShapeRange.Line
                                .EndArrowheadStyle = msoArrowheadOpen
                                .Visible = msoTrue
                                .ForeColor.RGB = RGB(255, 0, 0)
                                .Transparency = 0
                                .Weight = 2.5
                            End With
                        End With
                        
                        MsgBox ("Amortization value is not numeric, please make sure it's a number")
                        
                        Dim shp As Shape
                        For Each shp In ActiveSheet.Shapes
                           'shp.Delete
                        Next shp
                            
                        End
                    End If
                End If
            End If
        End With
    Next

'Nombrar Tab
'CrN = WB26.Worksheets("CECE").Cells(10, 3).Value

Dim sht As Worksheet
For Each sht In WB26.Worksheets
    If Application.Proper(sht.Name) = "Amortization" Then
        CrN = WB26.Worksheets("CECE").Cells(10, 3).Value
        WB26.Worksheets("Amortization").Name = CrN & " Amort"
        
        
    End If

Next sht


a = 1

End Sub


Sub ChecarWSht(shtName, wb)
    On Error GoTo EndNow
    Valor = Workbooks(wb).Worksheets("Pivot Summary").Range("A1").Value
    
EndNow:
    a = Err.Description
    If Err.Description = "Subscript out of range" Then
        Workbooks(wb).Sheets.Add.Name = "Pivot Summary"
    End If
    
    
End Sub



