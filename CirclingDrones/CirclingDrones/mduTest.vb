'*******************************************************
'      単体テスト用モジュール
'*******************************************************
Module mduTest

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 経路計算テスト
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subTest()

        Dim m As Integer = 1
        Dim n As Integer = 1
        pSet_d = 1

        '周回用ファイルを削除する
        Call subFileDel_w(pPjPath & "\d*.*")

        '***************************************************
        'New_Itiファイルを削除
        If fncFileDel(pExePath & "\New_Iti.csv") Then
            'fncMsgBox("New_Itiファイル削除=>OK")
        Else
            fncErrors("New_Itiファイル削除=>NG")
        End If
        '***************************************************
        '経路計算(n=0)する
        If fncItiCalc(m, 0) Then
            'fncMsgBox("経路計算=>OK")
        Else
            fncErrors("経路計算=>NG")
        End If

        '130.CSVファイルから130配列にセット
        If fncFile2Grid(pc130_Grid) Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        '130配列から飛距離計算する
        If fnc130Calc() Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        'Itiファイル(n=0)を作成する
        If fncItiFile(m, 0) Then
            'fncMsgBox("d" & m & "_Itiファイル作成=>OK")
        Else
            fncErrors("d" & m & "_Itiファイル作成=>NG")
        End If

        '***************************************************
        '経路計算する
        If fncItiCalc(m, n) Then
            'fncMsgBox("経路計算=>OK")
        Else
            fncErrors("経路計算=>NG")
        End If

        '130.CSVファイルから130配列にセット
        If fncFile2Grid(pc130_Grid) Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        '130配列から飛距離計算する
        If fnc130Calc() Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        'Itiファイルを作成する
        If fncItiFile(m, n) Then
            'fncMsgBox("d" & m & "_Itiファイル作成=>OK")
        Else
            fncErrors("d" & m & "_Itiファイル作成=>NG")
        End If

        'New_Itiファイルを作成する
        If fncNewItiFile(m, n) Then
            'fncMsgBox("New_Itiファイル作成=>OK")
        Else
            fncErrors("New_Itiファイル作成=>NG")
        End If
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' プロジェクト名.txrxファイルを作成する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncMem2txrx_test() As Boolean
        On Error GoTo Error_Rtn
        fncMem2txrx_test = False
        Const pcTag_test As String =
"begin_<polygon> Untitled polygon receiver set
project_id 8
active
invisible
vertical_line no
CVxLength 10.00000
CVyLength 10.00000
CVzLength 10.00000
AutoPatternScale
ShowDescription yes
CVsVisible no
CVsThickness 3
begin_<location> 
begin_<reference> 
cartesian
longitude 0.000000000000000
latitude 0.000000000000000
visible no
terrain
end_<reference>
spacing 0.80000
offset 0.00000
begin_<face> 
nVertices 4
185.55890 81.23181 2.00000
185.55890 -9.23169 2.00000
238.20907 -9.23169 2.00000
238.00301 81.64394 2.00000
end_<face>
end_<location>
pattern_show_arrow no
pattern_show_as_sphere no
generate_p2p no
use_apg_acceleration no
is_transmitter no
is_receiver yes
begin_<receiver> 
begin_<pattern> 
antenna 0
waveform 2
rotation_x 0.00000
rotation_y 0.00000
rotation_z 0.00000
end_<pattern>
NoiseFigure 3.00000
end_<receiver>
powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
end_<polygon>
begin_<points> Untitled points transmitter set
project_id 9
active
vertical_line no
cube_size 0.80000
CVxLength 10.00000
CVyLength 10.00000
CVzLength 10.00000
AutoPatternScale
ShowDescription yes
CVsVisible no
CVsThickness 3
begin_<location> 
begin_<reference> 
cartesian
longitude 0.000000000000000
latitude 0.000000000000000
visible no
terrain
end_<reference>
nVertices 1
226.797179664673310 72.275023203515872 2.000000000000000
end_<location>
pattern_show_arrow no
pattern_show_as_sphere no
generate_p2p yes
use_apg_acceleration no
is_transmitter yes
is_receiver no
begin_<transmitter> 
begin_<pattern> 
antenna 0
waveform 2
rotation_x 0.00000
rotation_y 0.00000
rotation_z 0.00000
end_<pattern>
power -30.00000
end_<transmitter>
powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
end_<points>"

        Dim sFile = pPjPath & "\" & pPjName & ".txrx"   'txrxファイルパス
        Dim iLoop As Integer

        'txrxファイル作成
        Dim enc As System.Text.Encoding = New System.Text.UTF8Encoding(False)
        Dim oFileWrite As New System.IO.StreamWriter(sFile, False, enc)

        'タグ
        Dim pTag_Test As String = pcTag_test.Replace(vbLf, vbCrLf)

        oFileWrite.WriteLine(pTag_Test)

        'クローズ
        oFileWrite.Dispose()
        oFileWrite.Close()

        fncMem2txrx_test = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function
End Module
