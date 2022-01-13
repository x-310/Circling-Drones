'*******************************************************
'	タグ作成用モジュール
'*******************************************************
Module mduTag

	'*******************************************************
	'	routeタグテンプレート
	'	nVerticesの初期値：2、最後は改行すること
	'*******************************************************
	'Public Const pcTag_route As String =
	'"begin_<route> d0 Route
	'nVertices 0
	'end_<route>
	'"

	'Public Const pcTag_route As String =
	'"begin_<route> d0 Route
	'vertical_line yes
	'end_<route>
	'"
	'

	Public Const pcTag_route As String =
"begin_<route> another drone route tx
project_id 19
active
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
longitude 138.555298569435930
latitude 36.412307004603001
visible no
terrain
end_<reference>
spacing 1.00000
nVertices 2
end_<location>
pattern_show_arrow no
pattern_show_as_sphere no
generate_p2p no
use_apg_acceleration no
is_transmitter yes
is_receiver no
begin_<transmitter> 
begin_<pattern> 
antenna 2
waveform 0
rotation_x 0.00000
rotation_y 0.00000
rotation_z 0.00000
end_<pattern>
power 30.00000
end_<transmitter>
powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
end_<route>
"
	Public pTag_route() As String           'routeタグ

	'*******************************************************
	'	gridタグテンプレート
	'	固定値、最後は改行すること
	'*******************************************************
	Public pTag_grid As String =
"begin_<grid> Untitled Rx Grid (3)" & vbCrLf &
"project_id 13" & vbCrLf &
"active" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription yes" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible no" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"side1 1637.29858" & vbCrLf &
"side2 1427.89490" & vbCrLf &
"spacing 14.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1002.624640000000000 1.571070000000000 50.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 1" & vbCrLf &
"waveform -1" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y 0.00000" & vbCrLf &
"rotation_z 0.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"begin_<sbr> " & vbCrLf &
"bounding_box" & vbCrLf &
"end_<sbr>" & vbCrLf &
"NoiseFigure 3.00000" & vbCrLf &
"end_<receiver>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<grid>"

End Module
