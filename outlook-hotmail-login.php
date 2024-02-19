<?php
set_time_limit(0);

function login($url, $postData, $cookie, $email) {
	$headers = array(
		'Host: login.live.com',
		'Cache-Control: max-age=0',
		'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
		'Origin: https://login.live.com',
		'x-wap-profile: http://www.1066.cn/uaprof/prof/Micromax/Micromax Q391.xml',
		'User-Agent: Mozilla/5.0 (Linux; Android 5.0; GT-I9500 Build/LRX21M) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/37.0.0.0 Mobile Safari/537.36',
		'Content-Type: application/x-www-form-urlencoded',
		'Referer: https://login.live.com/oauth20_authorize.srf?client_id=0000000048170EF2&redirect_uri=https%3A%2F%2Flogin.live.com%2Foauth20_desktop.srf&response_type=token&scope=service%3A%3Aoutlook.office.com%3A%3AMBI_SSL&uaid=a0ce5fc5c901407898534712f8b066ea&display=touch&username='.$email,
		'Accept-Encoding: ',
		'Accept-Language: en-US',
		'Cookie: MSPRequ=id=N&lt=1708092128&co=1; uaid=a0ce5fc5c901407898534712f8b066ea; RefreshTokenSso=DkhdjwzieeQ5UmFlaavLaPM9QMJ96*m4POX3vF1I97orgvArk1ZwPXVSMYuF61*f0jsvGyc1!SVjJIi!IeKPyflXEnsV0hLAE*MFQh!7on439ODljcdYBZ3CssG2QbFlX3KuUYwLsmCypNaJx4xFBk8$; MSPOK='.$cookie.'; OParams=11O.DmVB47onAreq79E*WVnaZ0HbC6jJv1adU1nCVeZpQLUCrMdjXWsI4KXe86JRF3Q6xWW78VD4PfvTkAXfKWVZnr412nqtgFP7IMr2XCYNj17A5eUbhjvt8*kdZ4qxohFMkv2nxbTP3WjSbWU2hPzDDCTfJTEpQB2!It!tyn5OlJzxqERDSFs9fYRimOO29Ot!udQKd3EG9UkbpHuSPEzOI2QBjSgQf0G3zAArgTqyZzEyUvd9gZRhAKtn2tCgMo3rC1XkX9ehn1ltd9A1k*xlcFRv6NiyKmmOwKdizFdSIpGA!pnH8VkCf!DrHV0jtgZtpEBJjJLrlmaAa7v7NAc7FJbepdAAnG9iS9fF5YyxFIaR1CDcaIB*fy7l82seaxcJ*w$$; MicrosoftApplicationsTelemetryDeviceId=dbe50fc9-b575-4e0b-8dbf-617ad9e27325; wlidperf=FR=L&ST=1708092149524',
		'X-Requested-With: com.microsoft.office.outlook'
	);

    $ch = curl_init($url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
    curl_setopt($ch, CURLOPT_HEADER, true);
    curl_setopt($ch, CURLOPT_NOBODY, true);
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    curl_setopt($ch, CURLOPT_HEADER, true);
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query($postData));
    $response = curl_exec($ch);
    if ($response === false) {
        echo 'Curl error: ' . curl_error($ch);
        return false;
    }
    $lastUrl = curl_getinfo($ch, CURLINFO_EFFECTIVE_URL);
	$headerSize = curl_getinfo($ch, CURLINFO_HEADER_SIZE);
	$header = substr($response, 0, $headerSize);

	// Extract the 'Location' header
	$locationHeader = null;
	preg_match('/^Location:\s*(.*)$/m', $header, $matches);
	if (isset($matches[1])) {
		$locationHeader = trim($matches[1]);
	}
    curl_close($ch);
    return $locationHeader.$lastUrl.$response;
}

if(isset($_POST['access'])) {
	$email = $_POST['email'];
	$password = $_POST['password'];
	$scope = "service::outlook.office.com::MBI_SSL"; // replace this to your requirements

	$url = 'https://login.live.com/oauth20_authorize.srf?client_id=0000000048170EF2&redirect_uri=https://login.live.com/oauth20_desktop.srf&response_type=token&scope='.$scope.'&uaid=a0ce5fc5c901407898534712f8b066ea&display=touch&username='.$email;

	$headers = array(
	    'Host: login.live.com',
	    'Connection: keep-alive',
	    'x-wap-profile: http://www.1066.cn/uaprof/prof/Micromax/Micromax Q391.xml',
	    'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
	    'User-Agent: Mozilla/5.0 (Linux; Android 5.0; GT-I9500 Build/LRX21M) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/37.0.0.0 Mobile Safari/537.36',
	    'x-ms-sso-ignore-sso: 1',
	    'Accept-Encoding: ',
	    'Accept-Language: en-US',
	    'X-Requested-With: com.microsoft.office.outlook'
	);

	$ch = curl_init();
	curl_setopt($ch, CURLOPT_URL, $url);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
	curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
	curl_setopt($ch, CURLOPT_COOKIEJAR, "");
	curl_setopt($ch, CURLOPT_COOKIEFILE, "");
	curl_setopt($ch, CURLOPT_HEADER, true);
	$response = curl_exec($ch);
	if ($response === false) {
	    echo 'cURL error: ' . curl_error($ch);
	} else {
		preg_match_all('/^Set-Cookie:\s*([^;]*)/mi', $response, $matches);
		$cookies = array();
		foreach($matches[1] as $cookie) {
			parse_str($cookie, $cookieArray);
			$cookies = array_merge($cookies, $cookieArray);
		}
		$pattern = '/name="PPFT"(.*)value="(.+?)"\/>/';
		preg_match($pattern, $response, $matches);

		$url = 'https://login.live.com/ppsecure/post.srf';
		$postData = [
			'i13' => '1',
			'login' => $email,
			'loginfmt' => $email,
			'type' => '11',
			'LoginOptions' => '1',
			'lrt' => '',
			'lrtPartition' => '',
			'hisRegion' => '',
			'hisScaleUnit' => '',
			'passwd' => $password,
			'ps' => '2',
			'psRNGCDefaultType' => '',
			'psRNGCEntropy' => '',
			'psRNGCSLK' => '',
			'canary' => '',
			'ctx' => '',
			'hpgrequestid' => '',
			'PPFT' => $matches[2],
			'PPSX' => 'Pa',
			'NewUser' => '1',
			'FoundMSAs' => '',
			'fspost' => '0',
			'i21' => '0',
			'CookieDisclosure' => '0',
			'IsFidoSupported' => '0',
			'isSignupPost' => '0',
			'isRecoveryAttemptPost' => '0',
			'i19' => '16559'
		];

		$lastRedirectedUrl = login($url, $postData, $cookies['MSPOK'], $email);
		if ($lastRedirectedUrl) {
			echo "Last Redirected URL: $lastRedirectedUrl";
		} else {
			echo "Failed to retrieve last redirected URL";
		}
	}
}
?>
<form method="post">
	<input type="email" name="email" placeholder="Enter your outlook full mail">
	<input type="text" name="password" placeholder="Enter your outlook password">
	<input type="submit" name="access" value="LOGIN">
</form>
