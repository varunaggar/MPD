$svgPath = "/Users/varunaggarwal/downloads/UBS-VMC-6.svg"
$svgContent = Get-Content -Path $svgPath -Raw
$contentType = "image/svg+xml"
$uri = "https://www.digicert.com/services/v2/util/validate-vmc-logo"
$response = Invoke-WebRequest -Uri $uri -Method Put -Body $svgContent -ContentType $contentType -ErrorAction Stop
 
 
try{
$response = Invoke-WebRequest -Uri $uri -Method Put -Body $svgContent -ContentType $contentType -ErrorAction Stop
}
catch{
 
$streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close() 
}