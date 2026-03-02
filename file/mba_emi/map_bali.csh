#!/bin/csh

set fault1=fault/2017pusgen_subduksi2017.gmt
set fault2=fault/floresthrust.gmt
set fault3=fault/NT_Banda_fault.gmt
set fault4=fault/Sesar_Bali.gmt
set fault5=fault/sunda.dat
set trench=fault/trench.gmt
set volcano=volcano.txt
set eq1=eq_all.dat
set R=113.21/117.31/-12/-6.75
set line_cc=line.dat
gmt begin seismsitas_nov24 png
gmt set FONT_ANNOT 10
gmt set FONT 7,Helvetica-Bold,black
gmt set MAP_FRAME_TYPE=plain
	gmt set FORMAT_GEO_MAP ddd.x
	gmt makecpt -T-12000/4000/100 -Cterra -Z
	gmt grdimage @earth_relief_01s.grd -R$R -JM7i -C -Bxa0.5 -Bya0.5 -BWnSe+t -I+da
	gmt coast -W0.8p -Tdg116.6/-7.5+w0.7i+f+l,,,N -Lg116.7/-11.82+w100k+l+f
	gmt colorbar -Bxaf+l"Elevation (m)" -C -Dx13.2/1.5+w4c/0.3c+h+m
	echo 115.3 -7.0 "SEISMISITAS WILAYAH BALI DAN SEKITARNYA"|gmt pstext -F+f15p,Helvetica-Bold,black+jMC -Gwhite -t20 

	gmt psxy $trench -W2,black -Sf2/0.15i+l+t -Gblack 
	gmt psxy $fault4 -W0.8,black -t30
	gmt psxy $fault2 -W0.8,black -t30
	gmt plot $volcano -St0.35c -Gorange -W0.3,black -t30
	
	gmt psxy $line_cc -W2,black -t30 
	
	echo 115.21 -7.35 "A'" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.21 -11.4 "A" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	awk '{print $1,$2,$3,$4*0.12}' $eq1| gmt plot -W0.3,black -Sc -Ctabel.cpt -t30

	 
	#legend for bali region
	echo 115.0 -11.45 "Magnitudo" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
#	echo 115.4 -11.28 "M9" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
#	echo 115.1 -11.28 "M8" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.2 -11.73 "M7" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.0 -11.73 "M6" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC 
	echo 114.8 -11.73 "M5" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.2 -11.92 "M4" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.0 -11.92 "M3" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 114.8 -11.92 "M2" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC 
#	echo 113.0 -11.28 "M1" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC 

#	echo 115.1 -11.22 | gmt psxy -Gblack -Sc0.96c -W0.3p,black -t30 
	echo 115.2 -11.6 | gmt psxy -Gblack -Sc0.84c -W0.3p,black -t30 
	echo 115.0 -11.6 | gmt psxy -Gblack -Sc0.72c -W0.3p,black -t30 
	echo 114.8 -11.6 | gmt psxy -Gblack -Sc0.6c -W0.3p,black -t30 
	echo 115.2 -11.82 | gmt psxy -Gblack -Sc0.48c -W0.3p,black -t30
	echo 115.0 -11.82 | gmt psxy -Gblack -Sc0.36c -W0.3p,black -t30
	echo 114.8 -11.82 | gmt psxy -Gblack -Sc0.24c -W0.3p,black -t30
#	echo 113.0 -11.22 | gmt psxy -Gblack -Sc0.12c -W0.3p,black -t30

	echo 115.7 -11.46 "Depth (km)" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.7 -11.6 "< 60" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.7 -11.7 "60-300" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC 
	echo 115.7 -11.8 ">300" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.7 -11.9 "Volcano" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
	echo 115.5 -11.6 | gmt psxy -Gred -Sc0.36c -W0.3p,black -t30 
	echo 115.5 -11.7 | gmt psxy -Gyellow -Sc0.36c -W0.3p,black -t30 
	echo 115.5 -11.8 | gmt psxy -Ggreen -Sc0.36c -W0.3p,black -t30 
	echo 115.5 -11.9 | gmt psxy -Gorange -St0.36c -W0.3p,black -t30
    	# Cross-section plot for earthquake depth profile
	awk '{print $1, $2, $3}' $eq1 | project -Q -C115.21/-11.4 -E115.21/-7.2 -Fxyzpqrs -W-222/222 > eq_cross.dat
	gmt basemap -JX5c/-3c -R0/450/0/450 -Bx100+l"Jarak (km)" -By100+l"Kedalaman (km)" -BWSne -Y0.5c -X1.2c
    	echo 10 -10 "A" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
    	echo 470 -10 "A'" | gmt pstext -F+f8.5p,Helvetica-Bold,black+jMC
    	awk '{print $4, $3, $3}' eq_cross.dat | gmt plot -Sc0.2c -Ctabel.cpt -W0.3p,black -t30

    	# Track and plot slab profile
    	project -C115.21/-11.5 -E115.21/-7.2 -Q -G1 > track2
    	grdtrack track2 -Gsum_slab1.0_clip.grd > tracked2
    	awk '{print $3, $4*-1}' tracked2 | gmt plot -W2p,black 
#	echo 115.45 -8.62 21 185 48 -151 75 68 -45 4.8 0 115.5 -8.62 | gmt psmeca -Sc0.8c -W0.3,black -Gblack
#	echo 115.5 -8.45 17 82 57 -76 238 34 -109 4.8 0 115.5 -8.45 | gmt psmeca -Sc0.8c -W0.4,black -Gblack
	#lon lat depth str dip slip st dip slip mant exp plon plat

	gmt inset begin -Dx-0.05/10.8+w3.4c/4.2c+o0.1c -F+gwhite+p1p
		#for main
		#gmt coast -R115/130.5/4/22 -JM -EPH+gkhaki+p0.2p -A10000 -W1p
		#gmt coast -R115/130.5/4/22 -JM -EPH+gdarkseagreen3+p0.2p -A10000 -W1p
		
		#for results map
		gmt coast -R116/128/4/22 -JM -EPH+ggray+p0.2p -A10000
		
		echo 118 12 126 20 | gmt plot -Sr+s -W1p,blue
		gmt psxy $trench -W0.5,black -Sf0.4/0.025i+l+t  
		gmt psxy $fault3 -W0.5,black
		gmt psxy $fault2 -W0.5,black 
		gmt psxy $fault1 -W0.5,black -Sf1c/0.0001c -Gblack 
		gmt psxy $arr1 -W0.1,red -Gwhite 
		gmt psxy $arr2 -W0.1,red -Gwhite
		gmt psxy $arr3 -W0.1,red -Gwhite
		gmt psxy $arr4 -W0.1,black 
		gmt pstext $name11 -F+f4p,Helvetica-Bold,black+jMC

    	gmt inset end

gmt end show
	
	
	

