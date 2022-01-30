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
        Dim n As Integer = 2
        pSet_d = 1

        '周回用ファイルを削除する
        Call subFileDel_w(pPjPath & "\d*.*")

        'New_Itiファイルを削除
        fncFileDel(pExePath & "\New_Iti.csv")

        ReDim Preserve p130(-1)

        'Itiファイル(n=0)を作成
        ReDim Preserve pIti(pcMax_d, 2)   '位置情報
        ReDim Preserve p130(2)

        Dim sX As String = ""
        Dim sY As String = ""
        Dim sZ As String = ""

        Select Case m
            Case 1
                sX = frmMain.txtX_d1.Text
                sY = frmMain.txtY_d1.Text
                sZ = frmMain.txtZ_d1.Text
            Case 2
                sX = frmMain.txtX_d2.Text
                sY = frmMain.txtY_d2.Text
                sZ = frmMain.txtZ_d2.Text
            Case 3
                sX = frmMain.txtX_d3.Text
                sY = frmMain.txtY_d3.Text
                sZ = frmMain.txtZ_d3.Text
            Case 4
                sX = frmMain.txtX_d4.Text
                sY = frmMain.txtY_d4.Text
                sZ = frmMain.txtZ_d4.Text
            Case 5
                sX = frmMain.txtX_d5.Text
                sY = frmMain.txtY_d5.Text
                sZ = frmMain.txtZ_d5.Text
        End Select

        pIti(m, 0).sIdo = sX
        pIti(m, 0).sKeido = sY
        pIti(m, 0).sTakasa = sZ

        pIti(m, 1).sIdo = sX
        pIti(m, 1).sKeido = sY
        pIti(m, 1).sTakasa = sZ

        pIti(m, 2).sIdo = "111"
        pIti(m, 2).sKeido = "112"
        pIti(m, 2).sTakasa = "113"

        fncItiFile(m, 0)

        '130配列を作成
        p130(0).dCal1 = 0.0
        p130(0).dCal2 = CDbl(frmMain.txtV.Text)
        p130(0).sX = sX
        p130(0).sY = sY
        p130(0).sZ = sZ

        '130.CSVファイルから130配列にセット
        fncFile2Grid(pc130_Grid)

        '130配列から飛距離計算する
        fncCalc130()

        '経路計算する
        fnc130Calc(m, n)

        '130.CSVファイルから130配列にセット
        fncFile2Grid(pc130_Grid)

        '130配列から飛距離計算する
        fncCalc130()

        'Itiファイルを作成する
        fncItiFile(m, n)

        'New_Itiファイルを作成する
        fncNewItiFile(m, n)

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

        'txrxファイル作成
        'Dim enc As System.Text.Encoding = New System.Text.UTF8Encoding(False)
        Dim oFileWrite As New System.IO.StreamWriter(sFile, False, pEnc)

        'タグ
        Dim pTag_Test As String = pcTag_test.Replace(vbLf, vbCrLf)

        oFileWrite.WriteLine(pTag_Test)

        'クローズ
        oFileWrite.Dispose()
        oFileWrite.Close()

        fncMem2txrx_test = True

    End Function
End Module
