VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF1 
   Caption         =   "UserForm1"
   ClientHeight    =   3276
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6756
   OleObjectBlob   =   "UF1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim controlLink As Object
    Dim controlSite As Object
    Dim controlTable As Object
    Dim controlRange As Object
    Dim existTable As Object
    Dim cond As Boolean
    Dim hasTable As Boolean
    Dim tableName As String
    Dim parseData
    
    cond = True
    hasTable = False
    
    Set controlLink = UF1.Controls("TextBoxLink")
    Set controlSite = UF1.Controls("ComboBoxSite")
    Set controlTable = UF1.Controls("TextBoxTable")
    Set controlRange = UF1.Controls("TextBoxRange")

    If controlLink.Value = "" Then
        cond = False
    End If
    If controlSite.Value = "" Then
        cond = False
    End If
    If controlTable.Value = "" And controlRange.Value = "" Then
        cond = False
    End If
    
    If Not cond Then
        MsgBox ("Заполните обязательные поля")
        Exit Sub
    End If
    
    If controlTable.Value <> "" Then
        Dim i As Integer
        
        For i = 1 To ActiveSheet.ListObjects.Count
            If ActiveSheet.ListObjects.Item(i).name = controlTable.Value Then
                hasTable = True
                tableName = controlTable.Value
                Exit For
            End If
        Next i
    End If
    
    If Not hasTable Then
        If controlRange.Value = "" Then
            MsgBox ("Нет таблицы и не указана ячейка")
            Exit Sub
        End If
        If controlTable.Value <> "" Then
            tableName = controlTable.Value
        Else
            tableName = "Таблица" & Format(Now, "mm/dd/yyyy HH:mm:ss")
        End If
        Call TableManager.CreateTable(controlRange.Value, tableName)
    End If
    
    parseData = Parsers.SendParser(controlLink.Value, controlSite.Value)
    Call TableManager.addRow(tableName, parseData)
    
    Unload UF1
End Sub

Private Sub CommandButton2_Click()
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To ActiveSheet.ListObjects.Count
        Dim table As Object
        Set table = ActiveSheet.ListObjects.Item(i)
        If table.ListColumns.Item(4).name = "Ссылка" Then
            For j = 1 To table.ListRows.Count
                Call TableManager.updateRow(table.ListRows.Item(j))
            Next j
        End If
    Next i
    Unload UF1
End Sub

Private Sub UserForm_Initialize()
    Dim control As Object
    
    For Each control In UF1.Controls
        If TypeName(control) = "ComboBox" Then
            control.AddItem "White Goods"
        End If
    Next control
    
End Sub
