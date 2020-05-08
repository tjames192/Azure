function ConvertTo-Speech {
<#
.SYNOPSIS
Converts text to speech
 
.DESCRIPTION
Uses Azure speech service text-to-speech REST API and outputs a .wav file
 
.EXAMPLE
Get-UMMailbox | select -expand displayname | ConvertTo-Speech

Converts each displayname property value to an individual wav file named $displayname.wav

.LINK
https://docs.microsoft.com/en-us/azure/cognitive-services/speech-service/quickstart-python-text-to-speech
#>
	[cmdletBinding()]
	param (
		[parameter(ValueFromPipeline=$True)]
		[string[]]$tts
	)
	process {
		$access_token = $null
		$subscription_key = "YOUR_KEY_HERE"
		$fetch_token_url = 'https://eastus.api.cognitive.microsoft.com/sts/v1.0/issuetoken'

		$headers = @{
			'Ocp-Apim-Subscription-Key' = $subscription_key
		}

		$response = Invoke-RestMethod -Method POST -Uri $fetch_token_url -Headers $headers

		$access_token = $response

		$base_url = 'https://eastus.tts.speech.microsoft.com/'
		$path = 'cognitiveservices/v1'
		$constructed_url = $base_url + $path

		$headers = @{
			'Content-Type' = 'application/ssml+xml'
			'X-Microsoft-OutputFormat'= 'riff-24khz-16bit-mono-pcm'
			'X-Search-AppId' = (New-Guid | Select-Object -ExpandProperty Guid).replace('-','')
			'X-Search-ClientId' = (New-Guid | Select-Object -ExpandProperty Guid).replace('-','')
			'User-Agent' = 'PowerShellTextToSpeechApp'
			'Authorization' = 'Bearer ' + $access_token
		}

		$lang = 'en-US'
		$name = 'Microsoft Server Speech Text to Speech Voice (en-US, JessaNeural)'
		$Content = $tts

		$Body = '<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="'+$lang+'"><voice name="'+$name+'">'+$Content+'</voice></speak>'

		$Filename = "$content.wav"
		
		# if no sleep we will hit...
		# Invoke-RestMethod : The remote server returned an error: (429) Too Many Requests.
		start-sleep 7
    
    		$Response = Invoke-WebRequest -Uri $constructed_url -Method POST -Headers $headers -Body $Body -OutFile $Filename -passthru

    		# if request was successful return filename
    		if ($Response.statuscode -eq 200) {
      			$Filename
    		}
    		else {
      			Write-warning "ERROR: $Filename, HTTP: $Response.StatusDescription"
    		}
  	}
}
