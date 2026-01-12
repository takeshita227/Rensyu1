VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineStop 
   Caption         =   "ライン停止内容"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   OleObjectBlob   =   "LineStop.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LineStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StopClock As Boolean
Sub ShowLineStop()
    Set LineStopForm = New LineStop
    LineStopForm.Show
End Sub

Public Sub UpdateClock()
    NextClock = Now + TimeValue("00:01")
    Me.lblTime.Caption = Format(Now, "hh:mm")
    Application.OnTime NextClock, "CallClock"
End Sub
Private Sub UserForm_Initialize()
    txtDate.Value = Format(Date, "yyyy/mm/dd")
    txtDate.Enabled = False
    txtDuration.Enabled = False
    Dim i As Integer
    Dim currentHour As String
    Dim currentMinute As String
    
    
  '停止時間、開始時間のプルダウンの選択肢
    
    '  時の十の位（0～2）
    For i = 0 To 2
        Me.txtstopTime1.AddItem i
        Me.txtstartTime1.AddItem i
    Next i
    
    '  時の一の位（0～9）
    For i = 0 To 9
        Me.txtstopTime2.AddItem i
        Me.txtstartTime2.AddItem i
    Next i
    
    '  分の十の位（0～5）
    For i = 0 To 5
        Me.txtstopTime3.AddItem i
        Me.txtstartTime3.AddItem i
    Next i
    
    '  分の一の位（0～9）
    For i = 0 To 9
        Me.txtstopTime4.AddItem i
        Me.txtstartTime4.AddItem i
    Next i
    
    
   '入力部分の初期表示
   
    '  現在時刻を取得
    currentHour = Format(Now, "hh")   ' 時（2桁）
    currentMinute = Format(Now, "nn") ' 分（2桁）
    
    '  停止時間に現在時刻を設定
    Me.txtstopTime1.Value = Mid(currentHour, 1, 1)
    Me.txtstopTime2.Value = Mid(currentHour, 2, 1)
    Me.txtstopTime3.Value = Mid(currentMinute, 1, 1)
    Me.txtstopTime4.Value = Mid(currentMinute, 2, 1)
    
    '  開始時間も同じ（必要なら別の値に変更可能）
    Me.txtstartTime1.Value = Mid(currentHour, 1, 1)
    Me.txtstartTime2.Value = Mid(currentHour, 2, 1)
    Me.txtstartTime3.Value = Mid(currentMinute, 1, 1)
    Me.txtstartTime4.Value = Mid(currentMinute, 2, 1)
    
    ' 初期化時に設備名・対応を非表示
    comProcess1.Visible = False
    comProcess2.Visible = False
    Label9.Visible = False
    Label12.Visible = False

    ' 停止理由の初期化
    comReason.Clear
    comReason.AddItem "交換"
    comReason.AddItem "不具合"
    comReason.AddItem "切替・手直し"
    comReason.AddItem "計画休止"
    comReason.AddItem "その他"
    
    StopClock = False
    Call UpdateClock
       
End Sub

'停止時間項目に入力された4つの数字を一つの時間として格納しなおす

Private Function GetStopTimeFromCombo() As Date

    Dim stophour As Integer
    Dim stopminute As Integer
    
    ' 一つ目のプルダウンと二つ目のプルダウンを時とする
    stophour = Val(Me.txtstopTime1.Value & Me.txtstopTime2.Value)
    
    ' 三つ目のプルダウンと四つ目のプルダウンを分とする
    stopminute = Val(Me.txtstopTime3.Value & Me.txtstopTime4.Value)
    
    ' Date型に変換（秒は0固定）
    GetStopTimeFromCombo = TimeSerial(stophour, stopminute, 0)
End Function

'再開時間項目に入力された4つの数字を一つの時間として格納しなおす

Private Function GetstartTimeFromCombo() As Date

    Dim starthour As Integer
    Dim startminute As Integer
    
    ' 一つ目のプルダウンと二つ目のプルダウンを時とする
    starthour = Val(Me.txtstartTime1.Value & Me.txtstartTime2.Value)
    
    ' 三つ目のプルダウンと四つ目のプルダウンを分とする
    startminute = Val(Me.txtstartTime3.Value & Me.txtstartTime4.Value)
    
    ' Date型に変換（秒は0固定）
    GetstartTimeFromCombo = TimeSerial(starthour, startminute, 0)
End Function

'停止時間、再開時間から設備停止時間を計算する

Private Sub UpdateDuration()
    Dim stopT As Date, startT As Date, duration As Variant
    
    ' 値を取得
    stopT = GetStopTimeFromCombo()
    startT = GetstartTimeFromCombo()
    
    ' 入力チェック（再開時刻が停止時刻より後か）
    If startT <= stopT Then
        Me.txtDuration.Value = "時間エラー"
        Exit Sub
    End If
    
    ' 差分計算（再開 - 停止）
    duration = startT - stopT
    
    ' hh:mm形式で表示
    Me.txtDuration.Value = Format(duration, "hh:mm")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    StopClock_Force
End Sub

Private Sub txtID_Change()
    Dim rawID As String, ID As String
    Dim ws As Worksheet
    Set ws = Sheets("社員一覧")

    rawID = Trim(txtID.Text)
    If Len(rawID) < 8 Then Exit Sub

    ' 最新文字だけ残す
    ID = Mid(rawID, 5, 4)
    txtID.Text = ID

    ' 名前を反映
    Dim result As Variant
    On Error Resume Next
    result = Application.WorksheetFunction.VLookup(CLng(ID), ws.Range("A:B"), 2, False)
    On Error GoTo 0

    If Not IsError(result) And Not IsEmpty(result) Then
        txtName.Value = result
    Else
        txtName.Value = ""
    End If
End Sub

' Enterを完全無効化（フォーム送信はボタンのみ）
Private Sub txtID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub



Private Sub comReason_Change()
    ' 停止理由詳細のコンボボックス初期化
    comDetails.Clear
    comProcess1.Clear
    comProcess2.Clear

    ' 一旦非表示
    comProcess1.Visible = False
    comProcess2.Visible = False
    Label9.Visible = False
    Label12.Visible = False

    Select Case comReason.Value
                    
        Case "交換"
            comDetails.AddItem "工程or設備名A"
            comDetails.AddItem "工程or設備名B"
            comDetails.AddItem "工程or設備名C"
            comDetails.AddItem "工程or設備名D"
            
            ' 表示（詳細2）
            comProcess1.Visible = True
            comProcess2.Visible = True
            Label9.Visible = True
            Label12.Visible = True

            ' 詳細2の選択肢
            comProcess1.AddItem "研削治具"
            comProcess1.AddItem "組立治具"
            comProcess1.AddItem "砥石"
            comProcess1.AddItem "消耗部品"
    
        Case "不具合"
            comDetails.AddItem "調整"
            comDetails.AddItem "故障"

            ' 表示（詳細2）
            comProcess1.Visible = True
            comProcess2.Visible = True
            Label9.Visible = True
            Label12.Visible = True

            ' 詳細2の選択肢
            comProcess1.AddItem "外R(研削)"
            comProcess1.AddItem "内R(研削)"
            comProcess1.AddItem "組立"
            comProcess1.AddItem "組立(検査)"

        Case "切替・手直し"
            comDetails.AddItem "呼番切替"
            comDetails.AddItem "シリーズ切替"
            comDetails.AddItem "手直し"
            
        Case "その他"
            comDetails.AddItem "干渉手待ち"
            comDetails.AddItem "材料手待ち"
            comDetails.AddItem "ミーティング"
            comDetails.AddItem "朝礼"
            comDetails.AddItem "4S"

        Case Else
            comDetails.AddItem "　ー　"
    End Select
End Sub

Private Sub comProcess1_Change()

    Dim key As String
    Dim items As Variant

    comProcess2.Clear
    key = comProcess1.Value

    ' ▼マスター一覧（全パターンここにまとめる）
    Select Case key

        '===== 外R / 内R / 組立 / 組立(検査) =====
        Case "外R(研削)"
            items = Array("投入", "軌道研削", "超仕上げ", "単体洗浄", "単洗後脱油")

        Case "内R(研削)"
            items = Array("投入", "軌道研削", "内径研削", "A/G", "超仕上げ", "単体洗浄", "単洗後脱油")

        Case "組立"
            items = Array("R/M", "玉入れ", "玉割り", "保持器入れ", "保持器カシメ", "保持器カシメ後脱磁", _
                          "完成品洗浄", "完洗後脱油", "レーザー印字", "防錆", "外観整列")

        Case "組立(検査)"
            items = Array("トルク", "保持器", "ボールなし・リベットなし・異型番(シール溝ありなし)", _
                          "音", "すきま", "自動外観")

        '===== 研削治具 / 組立治具 / 砥石 / 消耗部品 =====
        Case "研削治具"
            items = Array("シュー", "B.P")

        Case "組立治具"
            items = Array("かしめ型", "スプリング", "フィーラー", "測定子", "トリメトロン", _
                          "ノズル", "フィルター", "音アーバー")

        Case "砥石(副資材?)"
            items = Array("砥石", "ダイヤ", "油", "グリース")

        Case "消耗部品"
            items = Array("クイル", "ベルト", "ベアリング", "砥石", "すきま", "自動外観")

        Case Else
            items = Array("該当なし")

    End Select

    ' ▼アイテムを追加
    Dim i As Long
    For i = LBound(items) To UBound(items)
        comProcess2.AddItem items(i)
    Next i

    comProcess2.Visible = True

End Sub


Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Dim NextRow As Long
    Dim msg As String
    Dim stopT As Date, startT As Date
    Dim frmLoading As Object

    stopT = GetStopTimeFromCombo()
    startT = GetstartTimeFromCombo()

    msg = ""
    
    ' ====== 未入力チェック ======
    If Trim(Me.txtDate.Value) = "" Then msg = msg & "・作成日" & vbCrLf
    If Trim(Me.txtstopTime1.Value) = "" Then msg = msg & "・停止時刻" & vbCrLf
    If Trim(Me.txtstopTime2.Value) = "" Then msg = msg & "・停止時刻" & vbCrLf
    If Trim(Me.txtstopTime3.Value) = "" Then msg = msg & "・停止時刻" & vbCrLf
    If Trim(Me.txtstopTime4.Value) = "" Then msg = msg & "・停止時刻" & vbCrLf
    If Trim(Me.txtstartTime1.Value) = "" Then msg = msg & "・再開時刻" & vbCrLf
    If Trim(Me.txtstartTime2.Value) = "" Then msg = msg & "・再開時刻" & vbCrLf
    If Trim(Me.txtstartTime3.Value) = "" Then msg = msg & "・再開時刻" & vbCrLf
    If Trim(Me.txtstartTime4.Value) = "" Then msg = msg & "・再開時刻" & vbCrLf
    If Trim(Me.txtDuration.Value) = "" Then msg = msg & "・停止時間" & vbCrLf
    If Trim(Me.txtID.Value) = "" Then msg = msg & "・担当者ID" & vbCrLf
    If Trim(Me.txtName.Value) = "" Then msg = msg & "・担当者名" & vbCrLf
    If Trim(Me.comReason.Value) = "" Then msg = msg & "・停止理由" & vbCrLf
    If Trim(Me.comDetails.Value) = "" Then msg = msg & "・停止理由詳細" & vbCrLf

    If msg <> "" Then
        MsgBox "以下の項目が未入力です：" & vbCrLf & msg, vbExclamation, "入力チェック"
        Exit Sub
    End If
   ' ====== 時間の前後チェック ======
    
    If startT <= stopT Then
        MsgBox "再開時刻は停止時刻より後の時間を入力してください。", vbExclamation, "時間エラー"
        Exit Sub
    End If
    ' ==============================

    ' ====== 「送信中…」表示 ======
   Set frmLoading = New frmLoading   ' ← 新しくフォームを作る
   frmLoading.Show vbModeless        ' ← モードレスで表示
   DoEvents                        ' ← 画面描画を待つ

    ' ==============================

    ' ====== 書き込み処理 ======
    Set ws = Sheets("ライン停止内容")

    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    With ws
        .Cells(NextRow, 1).Value = txtDate.Value
        .Cells(NextRow, 3).Value = Format(stopT, "hh:mm")
        .Cells(NextRow, 4).Value = Format(startT, "hh:mm")
        .Cells(NextRow, 5).Value = txtDuration.Value
        .Cells(NextRow, 6).Value = comReason.Value
        .Cells(NextRow, 7).Value = comDetails.Value
        .Cells(NextRow, 8).Value = comProcess1.Value
        .Cells(NextRow, 9).Value = comProcess2.Value
        .Cells(NextRow, 10).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
        .Cells(NextRow, 11).Value = Environ("ComputerName")
    End With

    Unload frmLoading
    MsgBox "送信完了！", vbInformation

    ThisWorkbook.Save
    Unload Me
End Sub

Private Sub com_OpenTenKey1_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.txtStopTime
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.txtStopTime.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey2_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.txtStartTime
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.txtStartTime.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub com_OpenTenKey3_Click()
    frmTenKeyTemplate.TextBox1.Text = ""
    
    ' txtDefective を直接渡す
    Set frmTenKeyTemplate.TargetTextBox = Me.txtID
    
    frmTenKeyTemplate.Show vbModal

    ' 値を戻す
    If frmTenKeyTemplate.EnteredValue <> "" Then
        Me.txtID.Text = frmTenKeyTemplate.EnteredValue
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Enterキーを制御（QR用以外は無効）
    If KeyCode = vbKeyReturn Then
        If Not (ActiveControl Is txtID Or ActiveControl Is txtYOBIBAN) Then
            KeyCode = 0
        End If
    End If
End Sub

Private Sub txtstopTime1_Change(): Call UpdateDuration: End Sub
Private Sub txtstopTime2_Change(): Call UpdateDuration: End Sub
Private Sub txtstopTime3_Change(): Call UpdateDuration: End Sub
Private Sub txtstopTime4_Change(): Call UpdateDuration: End Sub

Private Sub txtstartTime1_Change(): Call UpdateDuration: End Sub
Private Sub txtstartTime2_Change(): Call UpdateDuration: End Sub
Private Sub txtstartTime3_Change(): Call UpdateDuration: End Sub
Private Sub txtstartTime4_Change(): Call UpdateDuration: End Sub



