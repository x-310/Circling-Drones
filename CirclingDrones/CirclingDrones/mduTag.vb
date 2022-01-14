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
"begin_<route> another drone tx
project_id 2
active
vertical_line no
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
visible no
terrain
end_<reference>
spacing 1000.00000
nVertices 2
end_<location>
pattern_show_arrow no
pattern_show_as_sphere no
generate_p2p no
use_apg_acceleration yes
is_transmitter yes
is_receiver no
begin_<transmitter> 
begin_<pattern> 
antenna 3
waveform 3
rotation_x 0.00000
rotation_y 0.00000
rotation_z 0.00000
end_<pattern>
power 30.00000
end_<transmitter>
begin_<receiver> 
begin_<pattern> 
antenna 1
waveform -1
rotation_x 0.00000
rotation_y 0.00000
rotation_z 0.00000
end_<pattern>
begin_<sbr> 
bounding_box
end_<sbr>
NoiseFigure 3.00000
end_<receiver>
powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
end_<route>
"
	Public pTag_route() As String           'routeタグ

	'*******************************************************
	'	gridタグテンプレート
	'	固定値、最後は改行すること
	'*******************************************************
	Public pTag_grid As String =
"begin_<grid> Dense Pine 50m Rx Grid" & vbLf &
"project_id 6" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 50.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 15m Rx Grid" & vbLf &
"project_id 9" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 15.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 20m Rx Grid" & vbLf &
"project_id 10" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 20.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 25m Rx Grid" & vbLf &
"project_id 11" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 25.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 30m Rx Grid" & vbLf &
"project_id 12" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 30.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 35m Rx Grid" & vbLf &
"project_id 13" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 35.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 40m Rx Grid" & vbLf &
"project_id 14" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 40.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 45m Rx Grid" & vbLf &
"project_id 15" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 45.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<grid> Dense Pine 10m Rx Grid" & vbLf &
"project_id 23" & vbLf &
"active" & vbLf &
"vertical_line no" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible no" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"side1 467.07999" & vbLf &
"side2 259.95999" & vbLf &
"spacing 5.00000" & vbLf &
"nVertices 1" & vbLf &
"1079.332770000000000 537.803000000000000 10.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow no" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p no" & vbLf &
"use_apg_acceleration yes" & vbLf &
"is_transmitter no" & vbLf &
"is_receiver yes" & vbLf &
"begin_<receiver> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 7" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y 0.00000" & vbLf &
"rotation_z 0.00000" & vbLf &
"end_<pattern>" & vbLf &
"begin_<sbr> " & vbLf &
"bounding_box" & vbLf &
"end_<sbr>" & vbLf &
"NoiseFigure 3.00000" & vbLf &
"end_<receiver>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<grid>" & vbLf &
"begin_<points> Tx Points 10 shokusei 3" & vbLf &
"project_id 41" & vbLf &
"inactive" & vbLf &
"invisible" & vbLf &
"vertical_line yes" & vbLf &
"cube_size 25.00000" & vbLf &
"CVxLength 10.00000" & vbLf &
"CVyLength 10.00000" & vbLf &
"CVzLength 10.00000" & vbLf &
"AutoPatternScale" & vbLf &
"ShowDescription no" & vbLf &
"CVsVisible no" & vbLf &
"CVsThickness 3" & vbLf &
"begin_<location> " & vbLf &
"begin_<reference> " & vbLf &
"cartesian" & vbLf &
"longitude 138.544999999999990" & vbLf &
"latitude 36.402900000000002" & vbLf &
"visible yes" & vbLf &
"terrain" & vbLf &
"end_<reference>" & vbLf &
"nVertices 1" & vbLf &
"1450.247570000000000 561.187590000000000 1.000000000000000" & vbLf &
"end_<location>" & vbLf &
"pattern_show_arrow yes" & vbLf &
"pattern_show_as_sphere no" & vbLf &
"generate_p2p yes" & vbLf &
"use_apg_acceleration no" & vbLf &
"is_transmitter yes" & vbLf &
"is_receiver no" & vbLf &
"begin_<transmitter> " & vbLf &
"begin_<pattern> " & vbLf &
"antenna 3" & vbLf &
"waveform 3" & vbLf &
"rotation_x 0.00000" & vbLf &
"rotation_y -50.00000" & vbLf &
"rotation_z 90.00000" & vbLf &
"end_<pattern>" & vbLf &
"power 30.00000" & vbLf &
"end_<transmitter>" & vbLf &
"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbLf &
"end_<points>"

End Module
