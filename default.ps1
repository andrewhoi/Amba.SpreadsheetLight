Properties {
	$base_directory = Resolve-Path .
	$src_directory = "$base_directory\Source"
	$output_directory = "$base_directory\Build"
	$dist_directory = "$base_directory\Distribution"
	$xunit_path = "$src_directory\packages\xunit.runners.2.0.0-rc3-build2880\tools\xunit.console.exe"
	$mstest_path = "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe"
	$nuget_path = "$src_directory\.nuget\nuget.exe"
	$ilmerge_path = "$src_directory\packages\ILMerge.2.14.1208\tools\ILMerge.exe"
}

$nl = [Environment]::NewLine

FormatTaskName (("-"*25) + "[{0}]" + ("-"*25))

Task Default -Depends Clean, Test, NuGetPackage


Task Build -Depends Clean {	

	Write-Host "Building Amba.SpreadsheetLight.sln" -ForegroundColor Green

	Exec { msbuild "$src_directory\Amba.SpreadsheetLight.sln" /t:Build /p:Configuration=Release /v:quiet /p:OutDir=$output_directory } 
}

Task Clean {

	Write-Host "Creating BuildOutput directory" -ForegroundColor Green
	
	rmdir $dist_directory -ea SilentlyContinue -recurse
	rmdir $output_directory -ea SilentlyContinue -recurse

	Write-Host "Amba.SpreadsheetLight.sln" -ForegroundColor Green
	Exec { msbuild "$src_directory\Amba.SpreadsheetLight.sln" /t:Clean /p:Configuration=Release /v:quiet } 
}



task Test -depends Build {  

	Write-Host "Testing Amba.SpreadsheetLight.Test" -ForegroundColor Green
	
	$project = "Amba.SpreadsheetLight.Test"
		
	mkdir $output_directory\xunit\$project -ea SilentlyContinue | Out-Null
	
	.$xunit_path "$output_directory\$project.dll" -html "$output_directory\xunit\$project\index.html"
	
	Write-Host $nl
}

task NuGetPackage -depends Build, Merge {

	$packageVersion = "1.0.0"
	
	Write-Host ("Creating NuGet Package v{0}" -f $packageVersion) -ForegroundColor Green
	
	copy-item $src_directory\Amba.SpreadsheetLight.nuspec $dist_directory
	copy-item $output_directory\Amba.SpreadsheetLight.xml $dist_directory\lib\net40\
	
	exec { .$nuget_path pack $dist_directory\Amba.SpreadsheetLight.nuspec -BasePath $dist_directory -o $dist_directory -version $packageVersion }
}

task Merge -depends Build{
	Write-Host "Merging Amba.SpreadsheetLight.dll with DocumentFormat.OpenXml.dll v2.0" -ForegroundColor Green
	
	New-Item $dist_directory\lib\net40 -Type Directory | Out-Null
	$input_dlls = "$output_directory\Amba.SpreadsheetLight.dll $output_directory\DocumentFormat.OpenXml.dll" 
	
	Invoke-Expression "$ilmerge_path /targetplatform:v4 /internalize /allowDup /target:library /out:$dist_directory\lib\net40\Amba.SpreadsheetLight.dll $input_dlls"
	Write-Host "Finished."
	Write-Host $nl
}
