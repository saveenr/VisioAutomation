
Draw a Directed Graph from script

	Import-Module Visio

# We'll need these DLLs

	$VA_path = "C:\Users\Saveen\Documents\WindowsPowershell\Modules\Visio"
	$VA_dll = Join-Path $VA_path "VisioAutomation.dll"
	$VAS_dll = Join-Path $VA_path "VisioAutomation.Scripting.dll"
	$VA_asm = [System.Reflection.Assembly]::LoadFrom($VA_dll)
	$VAS_asm = [System.Reflection.Assembly]::LoadFrom($VAS_dll)

# Create the Graph

	$d = New-VisioDirectedGraph
	$n1 = $d.AddShape("1","Node1", "BASIC_U.VSS", "Rectangle")
	$n2 = $d.AddShape("2","Node2", "BASIC_U.VSS", "Rounded Rectangle")
	$d.AddConnection("3",$n1,$n2)

The below renders it via MSAGL

		$options = New-Object VisioAutomation.Models.DirectedGraph.MSAGLLayoutOptions
	New-VisioApplication
	New-VisioDocument
	$p = New-VisioPage
	$d.Render($p,$options)

The code below renders it via Visio's algorithms

	# Render
	$options = New-Object VisioAutomation.Models.DirectedGraph.VisioLayoutOptions
	New-VisioApplication
	New-VisioDocument
	$p = New-VisioPage
	$d.Render($p,$options)

# Draw a Directed Graph from XML

	$dg = Import-VisioModel c:\foo.xml
	$dg | Out-Visio

## Sample Directed Graph 1
	<directedgraph>
	  <page>
	    <renderoptions
	      usedynamicconnectors="true"
	      scalingfactor="20"
	    />
	    <shapes>
	      <shape id="n1" label="FOO1" stencil="server_u.vss" master="Server" url="http://microsoft.com" />
	      <shape id="n2" label="FOO2" stencil="server_u.vss" master="Email Server" url="http://contoso.com"/>
	      <shape id="n3" label="FOO3" stencil="server_u.vss" master="Proxy Server" url="\\isotope\public" />
	      <shape id="n4" label="FOO4" stencil="server_u.vss" master="Web Server">
	        <customprop name="prop1" value="value1"/>
	        <customprop name="prop2" value="value2"/>
	        <customprop name="prop3" value="value3"/>
	      </shape>
	      <shape id="n5" label="FOO4" stencil="server_u.vss" master="Application Server" />
	    </shapes>
	
	    <connectors>
	      <connector id="c1"  from="n1" to="n2" label="LABEL1" />
	      <connector id="c2" from="n2" to="n3" label="LABEL2" color="#ff0000" weight="2" />
	      <connector id="c3" from="n3" to="n4" label="LABEL1" color="#44ff00" />
	      <connector id="c4" from="n4" to="n5" label="" color="#0000ff" weight="5"/>
	      <connector id="c5" from="n4" to="n1" label="" />
	      <connector id="c6" from="n4" to="n3" label="" weight="10"/>
	    </connectors>
	
	  </page>
	
	</directedgraph>

## Sample Directed Graph 2
	<directedgraph> 
	  <page>
	    
	  <renderoptions
	      usedynamicconnectors="true"
	      scalingfactor="20"
	    />
	  <shapes>
	    <shape id="n1" label="8761|Gus" stencil="basflo_u.vss" master="Process" />
	    <shape id="n2" label="0|ProABCS" stencil="basflo_u.vss" master="Process" />
	    <shape id="n3" label="ABCS" stencil="basflo_u.vss" master="Stored data" />
	    <shape id="n4" label="Global Underwriting" stencil="basflo_u.vss" master="Stored data" />
	  </shapes>
	
	  <connectors>
	    <connector id="c1"  from="n1" to="n2" label="" />
	    <connector id="c2" from="n2" to="n3" label="" />
	    <connector id="c3" from="n4" to="n2" label="" />
	  </connectors>
	
	  </page>
	
	  <page>
	
	    <renderoptions
	      usedynamicconnectors="true"
	      scalingfactor="20"
	    />
	    <shapes>
	      <shape id="n1" label="A" stencil="basflo_u.vss" master="Process" />
	      <shape id="n2" label="B" stencil="basflo_u.vss" master="Process" />
	      <shape id="n3" label="C" stencil="basflo_u.vss" master="Stored data" />
	      <shape id="n4" label="D" stencil="basflo_u.vss" master="Stored data" />
	    </shapes>
	
	    <connectors>
	      <connector id="c1"  from="n1" to="n2" label="" />
	      <connector id="c2" from="n2" to="n3" label="" />
	      <connector id="c3" from="n4" to="n2" label="" />
	    </connectors>
	
	  </page>
	</directedgraph>

## Sample Directed Graph 3
	
	<directedgraph>
	    
	  <renderoptions
	    usedynamicconnectors="false"
	    scalingfactor="20"
	    />
	  <shapes>
	    <shape id="n1" label="8761|Gus" stencil="basflo_u.vss" master="Process" />
	    <shape id="n2" label="0|ProABCS" stencil="basflo_u.vss" master="Predefined process" />
	    <shape id="n3" label="ABCS" stencil="basflo_u.vss" master="Stored data" />
	    <shape id="n4" label="Global Underwriting" stencil="basflo_u.vss" master="Stored data" />
	  </shapes>
	
	  <connectors>
	    <connector id="c1"  from="n1" to="n2" label="1111" />
	    <connector id="c2" from="n2" to="n3" label="222222" />
	    <connector id="c3" from="n4" to="n2" label="3333333" />
	  </connectors>
	
	  </page>
	
	  <page>
	
	    <renderoptions
	      usedynamicconnectors="true"
	      scalingfactor="20"
	    />
	    <shapes>
	      <shape id="n1" label="A" stencil="basflo_u.vss" master="Process" />
	      <shape id="n2" label="B" stencil="basflo_u.vss" master="Predefined process" />
	      <shape id="n3" label="C" stencil="basflo_u.vss" master="Stored data" />
	      <shape id="n4" label="D" stencil="basflo_u.vss" master="Stored data" />
	    </shapes>
	
	    <connectors>
	      <connector id="c1"  from="n1" to="n2" label="" />
	      <connector id="c2" from="n2" to="n3" label="" />
	      <connector id="c3" from="n4" to="n2" label="" />
	    </connectors>
	
	  </page>
	</directedgraph>

## Sample Directed Graph 4

	<directedgraph>
	  <page>
	
	    <renderoptions
	        usedynamicconnectors="true"
	        scalingfactor="5"
	    />
	  <shapes>
	    <shape id="n1" label="Decision1" stencil="basflo_u.vss" master="Decision" />
	    <shape id="n2" label="Process2" stencil="basflo_u.vss" master="Process" />
	    <shape id="n3" label="Data3" stencil="basflo_u.vss" master="Data" />
	    <shape id="n4" label="Porcess4" stencil="basflo_u.vss" master="Process" />
	    <shape id="n5" label="Data5" stencil="basflo_u.vss" master="Data" />
	  </shapes>
	
	  <connectors>
	    <connector id="c1"  from="n1" to="n2" label="LABEL1" />
	    <connector id="c2" from="n2" to="n3" label="LABEL2" />
	    <connector id="c3" from="n3" to="n4" label="LABEL3" />
	    <connector id="c4" from="n3" to="n1" label=""/>
	    <connector id="c5" from="n4" to="n1" label=""/>
	    <connector id="c6" from="n4" to="n2" label=""/>
	    <connector id="c7" from="n1" to="n5" label="LABEL7"/>
	  </connectors>
	
	  </page>
	
	</directedgraph>
