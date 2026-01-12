VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CT 
   Caption         =   "C/T登録"
   ClientHeight    =   8745.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "CT.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.txtDate.Value = Format(Date, "yyyy/mm/dd")
    Me.txtTime.Value = Format(Now, "hh:mm")
    Me.txtDate.Enabled = False
    Me.txtTime.Enabled = False
    
End Sub

Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Dim NextRow As Long
    Dim waitForm As frmLoading
    
    ' ====== 「送信中…」表示 ======
    Set waitForm = New frmLoading
    waitForm.Show vbModeless
    DoEvents
    ' =============================
    
    Set ws = Sheets("CT登録")
    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    With ws
        .Cells(NextRow, 1).Value = txtDate.Value
        .Cells(NextRow, 2).Value = txtTime.Value
        .Cells(NextRow, 3).Value = ct_OR.Value
        .Cells(NextRow, 4).Value = ct_OI.Value
        .Cells(NextRow, 5).Value = ct_IR.Value
        .Cells(NextRow, 6).Value = ct_II.Value
        .Cells(NextRow, 7).Value = ct_OSF.Value
        .Cells(NextRow, 8).Value = ct_ISF.Value
        .Cells(NextRow, 9).Value = ct_KUMI.Value
        .Cells(NextRow, 10).Value = Format(Now, "yyyy/mm/dd hh:mm:ss") ' 送信日時
        .Cells(NextRow, 11).Value = Environ("ComputerName") ' 端末名
    End With

    ThisWorkbook.Save

    ' ====== 「送信中」フォームを閉じる ======
    Unload waitForm
    Set waitForm = Nothing
    ' ======================================

    MsgBox "送信完了！", vbInformation
    Unload Me
End Sub

Private Sub com_OpenTenKey1_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.ct_OR
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.ct_OR.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey2_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.ct_OI
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.ct_OI.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey3_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.ct_IR
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.ct_IR.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey4_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.ct_II
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.ct_II.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey5_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.ct_OSF
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.ct_OSF.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey6_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.ct_ISF
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.ct_ISF.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey7_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.ct_KUMI
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.ct_KUMI.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

