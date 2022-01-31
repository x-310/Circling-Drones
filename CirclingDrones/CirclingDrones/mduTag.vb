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

	Public Const pcTag_point As String =
"begin_<points> another drone tx 1
project_id 41
active
invisible
vertical_line yes
cube_size 25.00000
CVxLength 10.00000
CVyLength 10.00000
CVzLength 10.00000
AutoPatternScale
ShowDescription no
CVsVisible no
CVsThickness 3
begin_<location> 
begin_<reference> 
cartesian
longitude 138.544999999999990
latitude 36.402900000000002
visible yes
terrain
end_<reference>
nVertices 1
1
end_<location>
pattern_show_arrow yes
pattern_show_as_sphere no
generate_p2p yes
use_apg_acceleration no
is_transmitter yes
is_receiver no
begin_<transmitter> 
begin_<pattern> 
antenna 3
waveform 3
rotation_x 0.00000
rotation_y -50.00000
rotation_z 90.00000
end_<pattern>
power 30.00000
end_<transmitter>
powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
end_<points>"

	Public pTag_point() As String           'routeタグ

	'*******************************************************
	'	gridタグテンプレート
	'	固定値、最後は改行すること
	'*******************************************************
	Public pTag_grid As String =
"begin_<points> Untitled Tx Points" & vbCrLf &
"project_id 1" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 2" & vbCrLf &
"waveform -1" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 23.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<grid> Dense Pine 50m Rx Grid" & vbCrLf &
"project_id 6" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 50.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 15m Rx Grid" & vbCrLf &
"project_id 9" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 15.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 20m Rx Grid" & vbCrLf &
"project_id 10" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 20.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 25m Rx Grid" & vbCrLf &
"project_id 11" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 25.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 30m Rx Grid" & vbCrLf &
"project_id 12" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 30.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 35m Rx Grid" & vbCrLf &
"project_id 13" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 35.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 40m Rx Grid" & vbCrLf &
"project_id 14" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 40.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 45m Rx Grid" & vbCrLf &
"project_id 15" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 45.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<grid> Dense Pine 10m Rx Grid" & vbCrLf &
"project_id 23" & vbCrLf &
"active" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
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
"side1 467.07999" & vbCrLf &
"side2 259.95999" & vbCrLf &
"spacing 5.00000" & vbCrLf &
"nVertices 1" & vbCrLf &
"1079.332770000000000 537.803000000000000 10.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p no" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter no" & vbCrLf &
"is_receiver yes" & vbCrLf &
"begin_<receiver> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 7" & vbCrLf &
"waveform 3" & vbCrLf &
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
"end_<grid>" & vbCrLf &
"begin_<points> Untitled Tx Points2" & vbCrLf &
"project_id 24" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.000000000000000 180.000000000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 2" & vbCrLf &
"waveform -1" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 23.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> Untitled Tx Points 3" & vbCrLf &
"project_id 25" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1470.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 2" & vbCrLf &
"waveform -1" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 23.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> Untitled Tx Points 4" & vbCrLf &
"project_id 27" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1300.247570000000000 561.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 2" & vbCrLf &
"waveform -1" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -33.00000" & vbCrLf &
"rotation_z 90.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 23.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> Untitled Tx Points 5" & vbCrLf &
"project_id 28" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1450.247570000000000 561.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 2" & vbCrLf &
"waveform -1" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -50.00000" & vbCrLf &
"rotation_z 90.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 23.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> otameshiA" & vbCrLf &
"project_id 31" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line no" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
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
"longitude 138.561926518429660" & vbCrLf &
"latitude 36.409927283247356" & vbCrLf &
"visible no" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"0.000000000000000 0.000000000000000 2.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow no" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration yes" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 2" & vbCrLf &
"waveform -1" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y 0.00000" & vbCrLf &
"rotation_z 0.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 23.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point" & vbCrLf &
"project_id 32" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 3" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point2" & vbCrLf &
"project_id 33" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 8" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 43.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point3 Denseless" & vbCrLf &
"project_id 34" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 8" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 43.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point4 Denseless" & vbCrLf &
"project_id 35" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 3" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point5 hole Foli" & vbCrLf &
"project_id 36" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 3" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point6 around Foli" & vbCrLf &
"project_id 37" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 3" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point7 shokusei Foli" & vbCrLf &
"project_id 38" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 3" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> 30dBm tx Point8 shokusei Foli HAR" & vbCrLf &
"project_id 39" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1100.247570000000000 651.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 3" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -27.00000" & vbCrLf &
"rotation_z 23.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>" & vbCrLf &
"begin_<points> Tx Points 9 shokusei 2" & vbCrLf &
"project_id 40" & vbCrLf &
"inactive" & vbCrLf &
"invisible" & vbCrLf &
"vertical_line yes" & vbCrLf &
"cube_size 25.00000" & vbCrLf &
"CVxLength 10.00000" & vbCrLf &
"CVyLength 10.00000" & vbCrLf &
"CVzLength 10.00000" & vbCrLf &
"AutoPatternScale" & vbCrLf &
"ShowDescription no" & vbCrLf &
"CVsVisible no" & vbCrLf &
"CVsThickness 3" & vbCrLf &
"begin_<location> " & vbCrLf &
"begin_<reference> " & vbCrLf &
"cartesian" & vbCrLf &
"longitude 138.544999999999990" & vbCrLf &
"latitude 36.402900000000002" & vbCrLf &
"visible yes" & vbCrLf &
"terrain" & vbCrLf &
"end_<reference>" & vbCrLf &
"nVertices 1" & vbCrLf &
"1300.247570000000000 561.187590000000000 1.000000000000000" & vbCrLf &
"end_<location>" & vbCrLf &
"pattern_show_arrow yes" & vbCrLf &
"pattern_show_as_sphere no" & vbCrLf &
"generate_p2p yes" & vbCrLf &
"use_apg_acceleration no" & vbCrLf &
"is_transmitter yes" & vbCrLf &
"is_receiver no" & vbCrLf &
"begin_<transmitter> " & vbCrLf &
"begin_<pattern> " & vbCrLf &
"antenna 3" & vbCrLf &
"waveform 3" & vbCrLf &
"rotation_x 0.00000" & vbCrLf &
"rotation_y -33.00000" & vbCrLf &
"rotation_z 90.00000" & vbCrLf &
"end_<pattern>" & vbCrLf &
"power 30.00000" & vbCrLf &
"end_<transmitter>" & vbCrLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
"end_<points>"

End Module
