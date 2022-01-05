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
	'& vbCrLf &
	'"begin_<route> Untitled Rx Route" & vbCrLf &
	'"project_id 14" & vbCrLf &
	'"active" & vbCrLf &
	'"vertical_line no" & vbCrLf &
	'"CVxLength 10.00000" & vbCrLf &
	'"CVyLength 10.00000" & vbCrLf &
	'"CVzLength 10.00000" & vbCrLf &
	'"AutoPatternScale" & vbCrLf &
	'"ShowDescription no" & vbCrLf &
	'"CVsVisible no" & vbCrLf &
	'"CVsThickness 3" & vbCrLf &
	'"begin_<location> " & vbCrLf &
	'"begin_<reference> " & vbCrLf &
	'"cartesian" & vbCrLf &
	'"longitude 138.544999999999990" & vbCrLf &
	'"latitude 36.402900000000002" & vbCrLf &
	'"visible no" & vbCrLf &
	'"terrain" & vbCrLf &
	'"end_<reference>" & vbCrLf &
	'"spacing 5.00000" & vbCrLf &
	'"nVertices 5" & vbCrLf &
	'"1534.180670000000000 663.379180000000020 50.000000000000000" & vbCrLf &
	'"1353.565270000000100 556.761979999999990 50.000000000000000" & vbCrLf &
	'"1224.901069999999900 718.291280000000030 50.000000000000000" & vbCrLf &
	'"1412.162070000000100 777.594979999999960 50.000000000000000" & vbCrLf &
	'"1534.126569999999900 663.095479999999950 50.000000000000000" & vbCrLf &
	'"end_<location>" & vbCrLf &
	'"pattern_show_arrow no" & vbCrLf &
	'"pattern_show_as_sphere no" & vbCrLf &
	'"generate_p2p no" & vbCrLf &
	'"use_apg_acceleration yes" & vbCrLf &
	'"is_transmitter no" & vbCrLf &
	'"is_receiver yes" & vbCrLf &
	'"begin_<receiver> " & vbCrLf &
	'"begin_<pattern> " & vbCrLf &
	'"antenna 1" & vbCrLf &
	'"waveform -1" & vbCrLf &
	'"rotation_x 0.00000" & vbCrLf &
	'"rotation_y 0.00000" & vbCrLf &
	'"rotation_z 0.00000" & vbCrLf &
	'"end_<pattern>" & vbCrLf &
	'"begin_<sbr> " & vbCrLf &
	'"bounding_box" & vbCrLf &
	'"end_<sbr>" & vbCrLf &
	'"NoiseFigure 3.00000" & vbCrLf &
	'"end_<receiver>" & vbCrLf &
	'"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
	'"end_<route>" & vbCrLf &
	'"begin_<route> Untitled Rx Route (2)" & vbCrLf &
	'"project_id 15" & vbCrLf &
	'"active" & vbCrLf &
	'"invisible" & vbCrLf &
	'"vertical_line no" & vbCrLf &
	'"CVxLength 10.00000" & vbCrLf &
	'"CVyLength 10.00000" & vbCrLf &
	'"CVzLength 10.00000" & vbCrLf &
	'"AutoPatternScale" & vbCrLf &
	'"ShowDescription no" & vbCrLf &
	'"CVsVisible no" & vbCrLf &
	'"CVsThickness 3" & vbCrLf &
	'"begin_<location> " & vbCrLf &
	'"begin_<reference> " & vbCrLf &
	'"cartesian" & vbCrLf &
	'"longitude 138.544999999999990" & vbCrLf &
	'"latitude 36.402900000000002" & vbCrLf &
	'"visible no" & vbCrLf &
	'"terrain" & vbCrLf &
	'"end_<reference>" & vbCrLf &
	'"spacing 1.50000" & vbCrLf &
	'"nVertices 2" & vbCrLf &
	'"1534.107670000000100 662.965360000000030 0.000000000000000" & vbCrLf &
	'"1534.107670000000100 662.965360000000030 50.000000000000000" & vbCrLf &
	'"end_<location>" & vbCrLf &
	'"pattern_show_arrow no" & vbCrLf &
	'"pattern_show_as_sphere no" & vbCrLf &
	'"generate_p2p no" & vbCrLf &
	'"use_apg_acceleration yes" & vbCrLf &
	'"is_transmitter no" & vbCrLf &
	'"is_receiver yes" & vbCrLf &
	'"begin_<receiver> " & vbCrLf &
	'"begin_<pattern> " & vbCrLf &
	'"antenna 1" & vbCrLf &
	'"waveform -1" & vbCrLf &
	'"rotation_x 0.00000" & vbCrLf &
	'"rotation_y 0.00000" & vbCrLf &
	'"rotation_z 0.00000" & vbCrLf &
	'"end_<pattern>" & vbCrLf &
	'"begin_<sbr> " & vbCrLf &
	'"bounding_box" & vbCrLf &
	'"end_<sbr>" & vbCrLf &
	'"NoiseFigure 3.00000" & vbCrLf &
	'"end_<receiver>" & vbCrLf &
	'"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
	'"end_<route>" & vbCrLf &
	'"begin_<route> Untitled Rx Route (3)" & vbCrLf &
	'"project_id 16" & vbCrLf &
	'"active" & vbCrLf &
	'"vertical_line no" & vbCrLf &
	'"CVxLength 10.00000" & vbCrLf &
	'"CVyLength 10.00000" & vbCrLf &
	'"CVzLength 10.00000" & vbCrLf &
	'"AutoPatternScale" & vbCrLf &
	'"ShowDescription no" & vbCrLf &
	'"CVsVisible no" & vbCrLf &
	'"CVsThickness 3" & vbCrLf &
	'"begin_<location> " & vbCrLf &
	'"begin_<reference> " & vbCrLf &
	'"cartesian" & vbCrLf &
	'"longitude 138.544999999999990" & vbCrLf &
	'"latitude 36.402900000000002" & vbCrLf &
	'"visible no" & vbCrLf &
	'"terrain" & vbCrLf &
	'"end_<reference>" & vbCrLf &
	'"spacing 5.00000" & vbCrLf &
	'"nVertices 2" & vbCrLf &
	'"1533.966320000000000 663.723719999999960 50.000000000000000" & vbCrLf &
	'"1336.449720000000100 753.113219999999960 50.000000000000000" & vbCrLf &
	'"end_<location>" & vbCrLf &
	'"pattern_show_arrow no" & vbCrLf &
	'"pattern_show_as_sphere no" & vbCrLf &
	'"generate_p2p no" & vbCrLf &
	'"use_apg_acceleration yes" & vbCrLf &
	'"is_transmitter no" & vbCrLf &
	'"is_receiver yes" & vbCrLf &
	'"begin_<receiver> " & vbCrLf &
	'"begin_<pattern> " & vbCrLf &
	'"antenna 1" & vbCrLf &
	'"waveform -1" & vbCrLf &
	'"rotation_x 0.00000" & vbCrLf &
	'"rotation_y 0.00000" & vbCrLf &
	'"rotation_z 0.00000" & vbCrLf &
	'"end_<pattern>" & vbCrLf &
	'"begin_<sbr> " & vbCrLf &
	'"bounding_box" & vbCrLf &
	'"end_<sbr>" & vbCrLf &
	'"NoiseFigure 3.00000" & vbCrLf &
	'"end_<receiver>" & vbCrLf &
	'"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
	'"end_<route>" & vbCrLf &
	'"begin_<route> Untitled Rx Route (4)" & vbCrLf &
	'"project_id 17" & vbCrLf &
	'"active" & vbCrLf &
	'"invisible" & vbCrLf &
	'"vertical_line no" & vbCrLf &
	'"CVxLength 10.00000" & vbCrLf &
	'"CVyLength 10.00000" & vbCrLf &
	'"CVzLength 10.00000" & vbCrLf &
	'"AutoPatternScale" & vbCrLf &
	'"ShowDescription no" & vbCrLf &
	'"CVsVisible no" & vbCrLf &
	'"CVsThickness 3" & vbCrLf &
	'"begin_<location> " & vbCrLf &
	'"begin_<reference> " & vbCrLf &
	'"cartesian" & vbCrLf &
	'"longitude 138.544999999999990" & vbCrLf &
	'"latitude 36.402900000000002" & vbCrLf &
	'"visible no" & vbCrLf &
	'"terrain" & vbCrLf &
	'"end_<reference>" & vbCrLf &
	'"spacing 2.50000" & vbCrLf &
	'"nVertices 2" & vbCrLf &
	'"1533.641120000000000 663.521269999999960 0.000000000000000" & vbCrLf &
	'"1533.641120000000000 663.521269999999960 50.000000000000000" & vbCrLf &
	'"end_<location>" & vbCrLf &
	'"pattern_show_arrow no" & vbCrLf &
	'"pattern_show_as_sphere no" & vbCrLf &
	'"generate_p2p no" & vbCrLf &
	'"use_apg_acceleration yes" & vbCrLf &
	'"is_transmitter no" & vbCrLf &
	'"is_receiver yes" & vbCrLf &
	'"begin_<receiver> " & vbCrLf &
	'"begin_<pattern> " & vbCrLf &
	'"antenna 1" & vbCrLf &
	'"waveform -1" & vbCrLf &
	'"rotation_x 0.00000" & vbCrLf &
	'"rotation_y 0.00000" & vbCrLf &
	'"rotation_z 0.00000" & vbCrLf &
	'"end_<pattern>" & vbCrLf &
	'"begin_<sbr> " & vbCrLf &
	'"bounding_box" & vbCrLf &
	'"end_<sbr>" & vbCrLf &
	'"NoiseFigure 3.00000" & vbCrLf &
	'"end_<receiver>" & vbCrLf &
	'"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
	'"end_<route>" & vbCrLf &
	'"begin_<grid> Untitled Rx Grid" & vbCrLf &
	'"project_id 18" & vbCrLf &
	'"active" & vbCrLf &
	'"invisible" & vbCrLf &
	'"vertical_line no" & vbCrLf &
	'"CVxLength 10.00000" & vbCrLf &
	'"CVyLength 10.00000" & vbCrLf &
	'"CVzLength 10.00000" & vbCrLf &
	'"AutoPatternScale" & vbCrLf &
	'"ShowDescription no" & vbCrLf &
	'"CVsVisible no" & vbCrLf &
	'"CVsThickness 3" & vbCrLf &
	'"begin_<location> " & vbCrLf &
	'"begin_<reference> " & vbCrLf &
	'"cartesian" & vbCrLf &
	'"longitude 138.544999999999990" & vbCrLf &
	'"latitude 36.402900000000002" & vbCrLf &
	'"visible no" & vbCrLf &
	'"terrain" & vbCrLf &
	'"end_<reference>" & vbCrLf &
	'"side1 467.08350" & vbCrLf &
	'"side2 259.95938" & vbCrLf &
	'"spacing 5.00000" & vbCrLf &
	'"nVertices 1" & vbCrLf &
	'"1079.332770000000000 537.803000000000000 50.000000000000000" & vbCrLf &
	'"end_<location>" & vbCrLf &
	'"pattern_show_arrow no" & vbCrLf &
	'"pattern_show_as_sphere no" & vbCrLf &
	'"generate_p2p no" & vbCrLf &
	'"use_apg_acceleration yes" & vbCrLf &
	'"is_transmitter no" & vbCrLf &
	'"is_receiver yes" & vbCrLf &
	'"begin_<receiver> " & vbCrLf &
	'"begin_<pattern> " & vbCrLf &
	'"antenna 1" & vbCrLf &
	'"waveform -1" & vbCrLf &
	'"rotation_x 0.00000" & vbCrLf &
	'"rotation_y 0.00000" & vbCrLf &
	'"rotation_z 0.00000" & vbCrLf &
	'"end_<pattern>" & vbCrLf &
	'"begin_<sbr> " & vbCrLf &
	'"bounding_box" & vbCrLf &
	'"end_<sbr>" & vbCrLf &
	'"NoiseFigure 3.00000" & vbCrLf &
	'"end_<receiver>" & vbCrLf &
	'"powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10" & vbCrLf &
	'"end_<grid>"

	'*******************************************************
	'	サンプルデータ(test3_wet_foli.txrx)
	'*******************************************************
	'begin_<points> Untitled Tx Points
	'	project_id 1
	'	active
	'	vertical_line yes
	'	cube_size 25.00000
	'	CVxLength 10.00000
	'	CVyLength 10.00000
	'	CVzLength 10.00000
	'	AutoPatternScale
	'	ShowDescription no
	'	CVsVisible no
	'	CVsThickness 3
	'	begin_<location> 
	'		begin_<reference>
	'			cartesian
	'			longitude 138.544999999999990
	'			latitude 36.402900000000002
	'			visible yes
	'			terrain
	'		end_<reference>
	'		nVertices 1
	'		1100.247570000000000 651.187590000000000 1.000000000000000
	'	end_<location>
	'	pattern_show_arrow yes
	'	pattern_show_as_sphere no
	'	generate_p2p yes
	'	use_apg_acceleration no
	'	is_transmitter yes
	'	is_receiver no
	'	begin_<transmitter> 
	'		begin_<pattern>
	'			antenna 2
	'			waveform -1
	'			rotation_x 0.00000
	'			rotation_y -27.00000
	'			rotation_z 23.00000
	'		end_<pattern>
	'		power 23.00000
	'	end_<transmitter>
	'	powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
	'end_<points>
	''---------------------------------------------------------------------
	'begin_<route> Untitled Rx Route
	'	project_id 2
	'	inactive
	'	vertical_line no
	'	CVxLength 10.00000
	'	CVyLength 10.00000
	'	CVzLength 10.00000
	'	AutoPatternScale
	'	ShowDescription no
	'	CVsVisible no
	'	CVsThickness 3
	'		begin_<location>
	'			begin_<reference> 
	'				cartesian
	'				longitude 138.544999999999990
	'				latitude 36.402900000000002
	'				visible no
	'				terrain
	'			end_<reference>
	'			spacing 5.00000
	'			nVertices 5
	'			1534.180670000000000 663.379180000000020 50.000000000000000
	'			1353.565270000000100 556.761979999999990 50.000000000000000
	'			1224.901069999999900 718.291280000000030 50.000000000000000
	'			1412.162070000000100 777.594979999999960 50.000000000000000
	'			1534.126569999999900 663.095479999999950 50.000000000000000
	'		end_<location>
	'	pattern_show_arrow no
	'	pattern_show_as_sphere no
	'	generate_p2p no
	'	use_apg_acceleration yes
	'	is_transmitter no
	'	is_receiver yes
	'	begin_<receiver>
	'		begin_<pattern> 
	'			antenna 1
	'			waveform -1
	'			rotation_x 0.00000
	'			rotation_y 0.00000
	'			rotation_z 0.00000
	'		end_<pattern>
	'		begin_<sbr> 
	'			bounding_box
	'		end_<sbr>
	'		NoiseFigure 3.00000
	'	end_<receiver>
	'	powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
	'end_<route>
	''---------------------------------------------------------------------
	'begin_<route> Untitled Rx Route (2)
	'	project_id 3
	'	inactive
	'	invisible
	'	vertical_line no
	'	CVxLength 10.00000
	'	CVyLength 10.00000
	'	CVzLength 10.00000
	'	AutoPatternScale
	'	ShowDescription no
	'	CVsVisible no
	'	CVsThickness 3
	'	begin_<location> 
	'		begin_<reference>
	'			cartesian
	'			longitude 138.544999999999990
	'			latitude 36.402900000000002
	'			visible no
	'			terrain
	'		end_<reference>
	'		spacing 1.50000
	'		nVertices 2
	'		1534.107670000000100 662.965360000000030 0.000000000000000
	'		1534.107670000000100 662.965360000000030 50.000000000000000
	'	end_<location>
	'	pattern_show_arrow no
	'	pattern_show_as_sphere no
	'	generate_p2p no
	'	use_apg_acceleration yes
	'	is_transmitter no
	'	is_receiver yes
	'	begin_<receiver> 
	'		begin_<pattern>
	'			antenna 1
	'			waveform -1
	'			rotation_x 0.00000
	'			rotation_y 0.00000
	'			rotation_z 0.00000
	'		end_<pattern>
	'		begin_<sbr>
	'			bounding_box
	'		end_<sbr>
	'		NoiseFigure 3.00000
	'		end_<receiver>
	'	powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
	'end_<route>
	''---------------------------------------------------------------------
	'begin_<route> Untitled Rx Route (3)
	'	project_id 4
	'	inactive
	'	vertical_line no
	'	CVxLength 10.00000
	'	CVyLength 10.00000
	'	CVzLength 10.00000
	'	AutoPatternScale
	'	ShowDescription no
	'	CVsVisible no
	'	CVsThickness 3
	'	begin_<location> 
	'		begin_<reference>
	'			cartesian
	'			longitude 138.544999999999990
	'			latitude 36.402900000000002
	'			visible no
	'			terrain
	'		end_<reference>
	'	spacing 5.00000
	'	nVertices 2
	'	1533.966320000000000 663.723719999999960 50.000000000000000
	'	1336.449720000000100 753.113219999999960 50.000000000000000
	'	end_<location>
	'	pattern_show_arrow no
	'	pattern_show_as_sphere no
	'	generate_p2p no
	'	use_apg_acceleration yes
	'	is_transmitter no
	'	is_receiver yes
	'	begin_<receiver> 
	'		begin_<pattern>
	'		antenna 1
	'		waveform -1
	'		rotation_x 0.00000
	'		rotation_y 0.00000
	'		rotation_z 0.00000
	'		end_<pattern>
	'			begin_<sbr> 
	'			bounding_box
	'			end_<sbr>
	'		NoiseFigure 3.00000
	'	end_<receiver>
	'	powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
	'end_<route>
	''---------------------------------------------------------------------
	'begin_<grid> Dense Pine 50m Rx Grid
	'	project_id 6
	'	active
	'	invisible
	'	vertical_line no
	'	CVxLength 10.00000
	'	CVyLength 10.00000
	'	CVzLength 10.00000
	'	AutoPatternScale
	'	ShowDescription no
	'	CVsVisible no
	'	CVsThickness 3
	'	begin_<location> 
	'		begin_<reference>
	'			cartesian
	'			longitude 138.544999999999990
	'			latitude 36.402900000000002
	'			visible no
	'			terrain
	'		end_<reference>
	'		side1 467.08350
	'		side2 259.95938
	'		spacing 5.00000
	'		nVertices 1
	'		1079.332770000000000 537.803000000000000 50.000000000000000
	'	end_<location>
	'	pattern_show_arrow no
	'	pattern_show_as_sphere no
	'	generate_p2p no
	'	use_apg_acceleration yes
	'	is_transmitter no
	'	is_receiver yes
	'	begin_<receiver> 
	'		begin_<pattern>
	'			antenna 1
	'			waveform -1
	'			rotation_x 0.00000
	'			rotation_y 0.00000
	'			rotation_z 0.00000
	'		end_<pattern>
	'		begin_<sbr>
	'			bounding_box
	'		end_<sbr>
	'		NoiseFigure 3.00000
	'	end_<receiver>
	'	powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
	'end_<grid>
	''---------------------------------------------------------------------
	'begin_<grid> Untitled Rx Grid (3)
	'	project_id 7
	'	inactive
	'	vertical_line no
	'	CVxLength 10.00000
	'	CVyLength 10.00000
	'	CVzLength 10.00000
	'	AutoPatternScale
	'	ShowDescription yes
	'	CVsVisible no
	'	CVsThickness 3
	'	begin_<location> 
	'		begin_<reference>
	'			cartesian
	'			longitude 138.544999999999990
	'			latitude 36.402900000000002
	'			visible no
	'			terrain
	'		end_<reference>
	'		side1 1637.29858
	'		side2 1427.89490
	'		spacing 14.00000
	'		nVertices 1
	'		1002.624640000000000 1.571070000000000 50.000000000000000
	'	end_<location>
	'	pattern_show_arrow no
	'	pattern_show_as_sphere no
	'	generate_p2p no
	'	use_apg_acceleration yes
	'	is_transmitter no
	'	is_receiver yes
	'	begin_<receiver> 
	'		begin_<pattern>
	'			antenna 0
	'			waveform -1
	'			rotation_x 0.00000
	'			rotation_y 0.00000
	'			rotation_z 0.00000
	'		end_<pattern>
	'		begin_<sbr>
	'			bounding_box
	'		end_<sbr>
	'		NoiseFigure 3.00000
	'	end_<receiver>
	'	powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
	'end_<grid>
	''---------------------------------------------------------------------
	'begin_<grid> Dense Pine 10m Rx Grid
	'	project_id 8
	'	inactive
	'	invisible
	'	vertical_line no
	'	CVxLength 10.00000
	'	CVyLength 10.00000
	'	CVzLength 10.00000
	'	AutoPatternScale
	'	ShowDescription no
	'	CVsVisible no
	'	CVsThickness 3
	'	begin_<location> 
	'		begin_<reference>
	'			cartesian
	'			longitude 138.544999999999990
	'			latitude 36.402900000000002
	'			visible no
	'			terrain
	'		end_<reference>
	'		side1 10.00000
	'		side2 10.00000
	'		spacing 5.00000
	'		nVertices 1
	'		1079.332770000000000 537.803000000000000 10.000000000000000
	'	end_<location>
	'	pattern_show_arrow no
	'	pattern_show_as_sphere no
	'	generate_p2p no
	'	use_apg_acceleration yes
	'	is_transmitter no
	'	is_receiver yes
	'	begin_<receiver> 
	'		begin_<pattern>
	'			antenna 1
	'			waveform -1
	'			rotation_x 0.00000
	'			rotation_y 0.00000
	'			rotation_z 0.00000
	'		end_<pattern>
	'		begin_<sbr>
	'			bounding_box
	'		end_<sbr>
	'		NoiseFigure 3.00000
	'	end_<receiver>
	'	powerDistribution Uniform 10.00000 10.00000 inactive nosampling 10
	'end_<grid>
End Module
