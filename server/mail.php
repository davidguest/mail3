<?php

/*

David Guest 2015
IT Services
University of Sussex

A simple mail client to read and send email messages using
Exchange Web Services - compatible with Exchange Online and On premises installations

*/

include("ews.php");
ini_set("soap.wsdl_cache_enabled", "0");
header("Access-Control-Allow-Origin: *");
//error_reporting(0);



if(isset($_REQUEST["token"]) || (isset($_REQUEST["username"]) && isset($_REQUEST["password"]))) {
	
	if(isset($_REQUEST["token"])) {
	
		$token = str_replace(" ", "+", $_REQUEST["token"]);
		$tokenbase = mcrypt_decrypt(MCRYPT_RIJNDAEL_256, base64_encode(strrev(strtolower(date("MY,D", time())))) , base64_decode($token),  MCRYPT_MODE_CBC, md5(base64_encode(strrev(strtolower(date("MY,D", time()))))));
		$baseparts = explode("-||-", $tokenbase);
		
		//check that username and password is correct in LDAP
		$user = $baseparts[0];
		$pass = $baseparts[1];
		
	} else {
	
		$user = $_REQUEST["username"];
		$pass = $_REQUEST["password"];
	
	}
	
	//initialize with the authenticated login
	$mail = new Exchangeclient();
	$mail->init($user, $pass);
	
	$action = "token";
	if(isset($_REQUEST["action"])) {
		switch($_REQUEST["action"]) {
			case "inbox": $action= "inbox"; break;
			case "read": $action= "read"; break;
			case "send": $action= "send"; break;
			case "download": $action= "download"; break;
			case "data": $action= "data"; break;
		}
	}
	
	$count = @$mail->get_messagecount();
	
	
	if(!$count) {
	
		header('Content-Type: text/plain; charset=UTF-8');
		echo json_encode(array("error"=>"please check your login details"));
	
	} elseif($action=="token") {
	
		//encrypt the username and password so the password
		//doesn't have to be placed in plain text in the source code
		$fw = $user . "-||-" . $pass . "-||-" . rand(0,264);
		$dt =  base64_encode(mcrypt_encrypt (MCRYPT_RIJNDAEL_256, base64_encode(strrev(strtolower(date("MY,D", time())))) , $fw , MCRYPT_MODE_CBC, md5(base64_encode(strrev(strtolower(date("MY,D", time())))))));
		$response = array("loginToken"=>urlencode($dt));
		header('Content-Type: text/plain; charset=UTF-8');
		echo json_encode($response);
	
	} elseif($action=="inbox") {
		
		// return a list of messages
		//parameters - offset, number of messages to return
		$params = array("o"=>0, "q"=>20);
		foreach(array_keys($params) as $param) {
			if(isset($_REQUEST[$param])) {
				$params[$param] = intval($_REQUEST[$param]);
			}
		}
		if(isset($_REQUEST["p"])) {
			$params["o"] = $params["q"] * ($_REQUEST["p"]-1);
		}
		
		header('Content-Type: text/plain; charset=UTF-8');
		$output = @$mail->list_inbox($params["o"], $params["q"]);
		echo json_encode($output);

	} elseif($action=="read") {
		
		// return text of a particular message
		if(isset($_REQUEST["uid"])) {
		
			$uid = intval($_REQUEST["uid"]);
			
			$item = @$mail->list_inbox($uid, 1);
			$id = $item[0]["id"]; 
			$changekey = $item[0]["changekey"];
			$isread = $item[0]["isread"];
			$msg = @$mail->get_message($id);
			$msg["uid"] = $uid;
			$msg["mailboxtotal"] = number_format($count->TotalCount);
			if(!$isread) {
				$seen = @$mail->mark_as_read($id, $changekey);
				$msg["markedAsSeen"] = true;
			}
			
			header('Content-Type: text/plain; charset=UTF-8');
			echo json_encode($msg);
			
		} elseif(isset($_REQUEST["id"]) && isset($_REQUEST["changekey"])) {
		
			$id = str_replace(" ", "+", urldecode($_REQUEST["id"]));
			$changekey = str_replace(" ", "+", urldecode($_REQUEST["changekey"]));
			
			$msg = @$mail->get_message($id);
			$isread = $msg["isread"];
			if($isread==0) {
				$seen = @$mail->mark_as_read($id, $changekey);
				$msg["markedAsSeen"] = true;
			}
			$msg["mailboxtotal"] = number_format($count->TotalCount);
			
			header('Content-Type: text/plain; charset=UTF-8');
			echo json_encode($msg);
			
		} else {
		
			header('Content-Type: text/plain; charset=UTF-8');
			echo json_encode(array("error"=>"Please specify the UID of the message"));
		
		}

	} elseif($action=="send") {
	
		if(isset($_REQUEST["to"])) {
		
			$cc = $bcc = $subject = $body = '';
			$to = $_REQUEST["to"];
			$cc = @$_REQUEST["cc"];
			$bcc = @$_REQUEST["bcc"];
			$subject = @$_REQUEST["subject"];
			$body = @$_REQUEST["body"];
			$body = str_replace("----ampersand----", "&", $body);
			header('Content-Type: text/plain; charset=UTF-8');
			echo json_encode(@$mail->send_email($to, $subject, $body, $cc, $bcc));
		
		} else {
		
			header('Content-Type: text/plain; charset=UTF-8');
			echo json_encode(array("error"=>"Please specify at least one recipient"));
		
		}
	} elseif($action=="download") {
		$uid = intval($_REQUEST["uid"]);
		$seq = intval($_REQUEST["seq"]);
		
		$item = @$mail->list_inbox($uid, 1);
		$itemid = $item[0]["id"]; 
		$filelist = @$mail->get_attachment_list($itemid);
		if(isset($filelist[$seq])) {
			
			$filedetails = $filelist[$seq];
			$filename = $filedetails["filename"];
			$size = $filedetails["size"];
			$attachid = $filedetails["id"];
			
			$attachment = @$mail->get_attachment($attachid);
			
			
			header('Content-Disposition: attachment; filename="' . $filename . '"');
			header('Content-Type: ' . $attachment->ContentType);
			//header('Content-Length: ' . $size);
			echo @$mail->get_raw_attachment($attachid);
		
		} else {
		
			header('Content-Type: text/plain; charset=UTF-8');
			echo json_encode(array("error"=>"Could not locate an attachment for those details"));
		
		} 
			
		
		
	} elseif($action=="data") {
	
		$data = array();
		
		if (!empty($_SERVER['HTTP_X_FORWARDED_FOR'])) {
			$ip=$_SERVER['HTTP_X_FORWARDED_FOR'];
		} else {
			$ip=$_SERVER['REMOTE_ADDR'];
		}
		
		$data["all"] = number_format($count->TotalCount);
		$data["unseen"] = number_format($count->UnreadCount);
		$data["ip"] = $ip;
	
		header('Content-Type: text/plain; charset=UTF-8');
		echo json_encode($data);
	
	}

} else {

	header('Content-Type: text/plain; charset=UTF-8');
	echo json_encode(array("error"=>"Please specify login details"));

}


