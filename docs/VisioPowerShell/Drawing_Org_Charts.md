
#create an XML as shown below:

	$xmldoc = @"
	<orgchart>
	  <shape id="0" name="Akuma" />
	  <shape id="1" name="Ryu" parentid="0"/>
	  <shape id="2" name="Ken" parentid="0"/>
	  <shape id="3" name="Chun-Li" parentid="2"/>
	</orgchart>
	"@

	$xmldoc | out-file "d:\orgchart.xml"

Then load the XML and use Out-Visio

	$orgchart= Import-VisioModel -Filename D:\orgchart
