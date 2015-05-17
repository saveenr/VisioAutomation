# Charts

# Drawing a Pie Chart

	$centerx = 1
	$centery = 1
	$radius = 5
	$values = @(2,4,5,10)
	$chart1 = New-VisioPieChart $centerx $centery $radius $values
	$chart1 | Out-Visio

# Drawing a Doughnut Chart

Using -InnerRadius allows for a doughnut chart

	$centerx = 1
	$centery = 1
	$radius = 5
	$values = @(2,4,5,10)
	$inner_radius = 4
	$chart2 = New-VisioPieChart $centerx $centery $radius $values -InnerRadius $inner_radius
	$chart2 | Out-Visio

