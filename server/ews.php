<?php
/*
 * Exchange scripts calling EWS on Exchange Online
 * By David Guest for University of Sussex 2015
 *
 * based on publicly available classes from
 * http://www.howtoforge.com/talking-soap-with-exchange
 *
 */
 
 
class Exchangeclient {  

	private $wsdl;
	private $client;
	private $user;
	private $pass;
	private $version;
	/**
	 * The last error that occurred when communicating with the Exchange server.
	 * 
	 * @var mixed
	 * @access public
	 */
	public $lastError;
	private $impersonate;
	
	
		

	/**
	 * Initialize the class. 
	 * 
	 * @access public
	 * @param string $user (the username of the mailbox account you want to use on the Exchange server)
	 * @param string $pass (the password of the account)
	 * @param string $impersonate. (the email address of someone you want to impoersonate on the server [NOTE: Your server must have Impersonation enabled (most don't) and you're account must have impersonation permissions (most don't), otherwise this will cause everything to fail! If you don't know what it is, leave it alone :-)] default: NULL)
	 * @param string $wsdl. (The path to the WSDL file.)
	 * @return void
	 */
	function init($user, $pass, $impersonate=NULL, $wsdl="Services.wsdl", $version="Exchange2013_SP1") {
	
		//$wsdl = "https://outlook.office365.com/EWS/Services.wsdl";
		$this->wsdl = "wsdl/$wsdl";
		$this->user = $user;
		$this->pass = $pass;
		$this->version = $version;

		$this->client = new ExchangeNTLMSoapClient($this->wsdl);
		$this->client->user = $this->user;
		$this->client->password = $this->pass;
		
		if($impersonate != NULL) {
			$this->impersonate = $impersonate;
		}

	}
   

	/*
		Get time zones
		by David Guest
	*/
	
	function get_timezones($timeZone="") {
		$this->setup();
		
		if($timeZone!="") {
			$GetServerTimeZones->ReturnFullTimeZoneData = "true";
			$GetServerTimeZones->Ids->Id = $timeZone;
			
		} else {
			$GetServerTimeZones->ReturnFullTimeZoneData = "false";
		}
		
		$response = $this->client->GetServerTimeZones($GetServerTimeZones);

		$this->teardown();

		if($response->ResponseMessages->GetServerTimeZonesResponseMessage->ResponseCode == "NoError") {
		
			$zones = $response->ResponseMessages->GetServerTimeZonesResponseMessage->TimeZoneDefinitions->TimeZoneDefinition;
			return $zones;
		} else {
		
			return $response;
		}
	
	}
	
	/*
		Get message count
		by David Guest
	*/
	
	function get_messagecount() {
		$this->setup();
		
		$GetFolder->FolderShape->BaseShape = "Default";
		$GetFolder->FolderIds->DistinguishedFolderId->Id = "inbox";
		$response = $this->client->GetFolder($GetFolder);
		
		$this->teardown();
		
		$folder = $response->ResponseMessages->GetFolderResponseMessage->Folders->Folder;
		
		return $folder;
		
	}
	
	/*
		list inbox messages
		by David Guest
	*/
	
	
	function list_inbox($offset="0", $count="10") {
	
		$this->setup();
		
		$FindItem->Traversal="Shallow";
		$FindItem->ItemShape->BaseShape="AllProperties";
		$FindItem->ParentFolderIds->DistinguishedFolderId->Id = "inbox";
		$FindItem->IndexedPageItemView->BasePoint = "Beginning";
		$FindItem->IndexedPageItemView->Offset = $offset;
		$FindItem->IndexedPageItemView->MaxEntriesReturned = $count;
		$response = $this->client->FindItem($FindItem);
		
		$itemtypes = array();
		$items = $response->ResponseMessages->FindItemResponseMessage->RootFolder->Items;
		foreach($items as $itemtype) {
			$itemtypes[] = $itemtype;
		}
		
		$this->teardown();

		$output = array();
		foreach($itemtypes as $messages) {
			if(!is_array($messages)) {
				$messages = array($messages);
			}
			foreach($messages as $message) {
				$unixtime = strtotime(@$message->DateTimeReceived);
				$stamp = date("j M H:i", $unixtime);
				$attachmentcount = intval(@$message->HasAttachments);
				$from = array("name"=>@$message->Sender->Mailbox->Name);
				$newitem = array(
									"subject"=>@$message->Subject, 
									"from"=>@$from,
									"attachments"=>@$attachmentcount,
									"date"=>@$stamp,
									"id"=>@$message->ItemId->Id,
									"changekey"=>@$message->ItemId->ChangeKey,
									"isread"=>@$message->IsRead,
									"unixtime"=>$unixtime
								);
				if(count($output)<=0) {
					$output[] = $newitem;
				} else {
					//iterate through and put them back in datetime order
					$m = 0; $lastm = count($output)-1;
					foreach($output as $thisitem) {
						if($newitem["unixtime"] > $thisitem["unixtime"]) {
							$insert = array($newitem);
							$head = array_slice($output, 0, $m);
							$tail = array_slice($output, $m);
							$output = array_merge($head, $insert, $tail);
							break;
						} elseif($m==$lastm) {
							$output[] = $newitem;
						}
						$m++;
					}
				}
			}
		}
		for($o=0;$o<count($output);$o++) {
			$output[$o]["uid"] = $offset+$o;
		}
		return $output;
	}
	
	/*
		Get message
		by David Guest
	*/
	
	function get_message($itemid) {
		$this->setup();
		
		$GetItem->ItemShape->BaseShape = "Default";
		$GetItem->ItemShape->BodyType = "Text";
		$GetItem->ItemIds->ItemId->Id = $itemid;
		
		$response = $this->client->GetItem($GetItem); 

		$this->teardown();
		
		//find the address of the current user
		$testfor = $this->user;
		if(stristr($testfor, '\\')) {
			$bits = explode('\\', $testfor);
			$testfor = $bits[1];
		} 
		if(stristr($testfor, "@")) {
			$bits = explode("@", $testfor);
			$testfor = $bits[0];
		}
		$resolve = $this->resolve_name($testfor);
		$current_email = @$resolve->ResolutionSet->Resolution->Mailbox->EmailAddress;
		
		if($response->ResponseMessages->GetItemResponseMessage->ResponseCode == "NoError") {
		
			$items = $response->ResponseMessages->GetItemResponseMessage->Items;
			foreach($items as $item) {
				$message = $item;
			}
			if(isset($items->Message)) {
				$itemtype = "email";
			} elseif(isset($items->CalendarItem) || isset($items->MeetingMessage) || isset($items->MeetingRequest) || isset($items->MeetingResponse) || isset($items->MeetingCancellation)) {
				$itemtype = "calendar";
			} else {
				$itemtype = "other";
			}
			$output = array(); 
			$output["subject"] = $message->Subject;
			if(isset($message->From)) {
				$frommail = @$message->From->Mailbox->EmailAddress;
				$fromname = @$message->From->Mailbox->Name;
			} elseif(isset($message->Organizer)) {
				$frommail = @$message->Organizer->Mailbox->EmailAddress;
				$fromname = @$message->Organizer->Mailbox->Name;
			} else {
				$frommail = $fromname = "";
			}
			$output["from"] = ["name"=>$fromname, "address"=>$frommail];
			
			//reorganise recipients
			$to = $message->ToRecipients->Mailbox;
			$torecipients = [$frommail];
			$sentto = array();
			if(!is_array($to)) { $to = [$to]; }
			foreach($to as $toperson) {
				$tomail = $toperson->EmailAddress;
				$sentto[] = $tomail;
				if($tomail != $current_email) { $torecipients[] = $tomail; }
			}
			$output["recipients"]["gobackto"] = implode(", ", $torecipients);
			$output["recipients"]["to"] = implode(", ", $sentto);
			
			$cc = $message->CcRecipients->Mailbox;
			$ccrecipients = array();
			if(!is_array($cc)) { $cc = [$cc]; }
			foreach($cc as $ccperson) {
				$ccrecipients[] = $ccperson->EmailAddress;
			}
			$output["recipients"]["cc"] = implode(", ", $ccrecipients);
			
			//date and time stamp
			if(isset($message->DateTimeSent)) {
				$timestmp = strtotime($message->DateTimeSent);
				$output["date"] = date("j M H:i", $timestmp);
			} elseif(isset($message->Start)) {
				$starts = strtotime($message->Start);
				$ends = strtotime($message->End);
				$output["date"] = date("j M H:i", $starts) . " - " . date("j M H:i", $ends);
			} else {
				$output["date"] = "";
			}
			
			//check if this item has been read
			$output["isread"] = $message->IsRead;
			
			//message body
			$rawbody = $message->Body->_;
			
			//try to cut down on white space
			$rawbody = str_replace("\r\n\r\n\r\n","\r\n\r\n",$rawbody);
			$rawbody = str_replace("\r\n\r\n\r\n","\r\n\r\n",$rawbody);
			$output["bodytext"] = $rawbody;
			
			//parse links in text
			$links = array();
			$rawbody .= " ";
			$rawbody = str_replace("&lt;", "<", $rawbody);
			$rawbody = str_replace("&gt;", ">", $rawbody);
			$formatted_links = $this->scrape_all($rawbody,"<http",">");
			foreach($formatted_links as $formatted_link) {
				$candidatelink = "http" . $formatted_link;
				if(!in_array($candidatelink, $links)) {
					$links[] = $candidatelink;
				}
			}
			$replace_these = array("-\r\n", "?\r\n", "\r","\n",">",")","]");
			foreach($replace_these as $replace_this) {
				$rawbody = str_replace($replace_this, " ", $rawbody);
			}
			$http_links = $this->scrape_all($rawbody, "http://"," ");
			foreach($http_links as $http_link) {
				$candidatelink = "http://" . rtrim($http_link, ']>.)"');
				if(!in_array($candidatelink, $links)) {
					$links[] = $candidatelink;
				}
			}
			$https_links = $this->scrape_all($rawbody, "https://"," ");
			foreach($https_links as $https_link) {
				$candidatelink = "https://" . rtrim($https_link, ']>.)"');
				if(!in_array($candidatelink, $links)) {
					$links[] = $candidatelink;
				}
			}
			//sort the links so longest are first
			usort($links, function($a, $b) {
				return strlen($b) - strlen($a);
			});
			$output["links"] = $links;
			
			//create a body if this is a calendar or task item
			if($itemtype == "calendar" || $itemtype=="other") {
				$output["bodytext"] = "You can manage this item using the Outlook web app.";
			} 
			
			//attachments
			$attachments = $this->get_attachment_list($itemid);
			$output["attachments"] = $attachments;
			return $output;
			
		} else {
		
			return false;
		}
	
	}
	
	/*
		Mark as read
		by David Guest
	*/
	
	function mark_as_read($itemid, $changekey) {
		$this->setup();
		
		$UpdateItem->MessageDisposition = "SaveOnly";
		$UpdateItem->ConflictResolution = "AutoResolve";
		$UpdateItem->ItemChanges->ItemChange->ItemId->Id = $itemid;
		$UpdateItem->ItemChanges->ItemChange->ItemId->ChangeKey = $changekey;
		$UpdateItem->ItemChanges->ItemChange->Updates->SetItemField->FieldURI->FieldURI = "message:IsRead";
		$UpdateItem->ItemChanges->ItemChange->Updates->SetItemField->Message->IsRead = true;
		$response = $this->client->UpdateItem($UpdateItem);

		$this->teardown();

		return $response;
	
	}
	
	/*
		Get raw attachment
		by David Guest
	*/
	
	function get_raw_attachment($attachid) {
		$this->setup();
		
		$GetAttachment->AttachmentShape->IncludeMimeContent = true;
		//$GetAttachment->AttachmentShape->BodyType = "";
		$GetAttachment->AttachmentIds->AttachmentId->Id = $attachid;
		$response = $this->client->GetAttachment($GetAttachment);

		$this->teardown();

		$attachment = $response->ResponseMessages->GetAttachmentResponseMessage->Attachments->FileAttachment;
		return $attachment->Content;
	
	}
	
	/*
		Get attachment details
		by David Guest
	*/
	
	function get_attachment($attachid) {
		$this->setup();
		
		$GetAttachment->AttachmentShape->IncludeMimeContent = true;
		//$GetAttachment->AttachmentShape->BodyType = "";
		$GetAttachment->AttachmentIds->AttachmentId->Id = $attachid;
		$response = $this->client->GetAttachment($GetAttachment);

		$this->teardown();
		
		$attachment = $response->ResponseMessages->GetAttachmentResponseMessage->Attachments->FileAttachment;
		if(!isset($attachment->ContentType)) {
			$attachment->ContentType = "text/plain";
		}
		unset($attachment->Content);
		return $attachment;
	
	}
	
	/*
		Get attachment list
		by David Guest
	*/
	
	function get_attachment_list($itemid) {
	
		$this->setup();
		
		$GetItem->ItemShape->BaseShape = "Default";
		$GetItem->ItemShape->BodyType = "Text";
		$GetItem->ItemIds->ItemId->Id = $itemid;
		
		$response = $this->client->GetItem($GetItem); 

		$this->teardown();
		
		$attachments = array();
		if($response->ResponseMessages->GetItemResponseMessage->ResponseCode == "NoError") {
			$message = $response->ResponseMessages->GetItemResponseMessage->Items->Message;
			if(isset($message->Attachments)) {
				$files = $message->Attachments->FileAttachment;
				$files = !is_array($files) ? array($files) : $files;
				$seq = 0;
				foreach($files as $file) {
					if(intval($file->Size) > 0) {
						$attachments[] = array("filename"=>$file->Name, "id"=>$file->AttachmentId->Id, "seq"=>$seq, "size"=>$file->Size);
						$seq++;
					}
				}
			} 
		}
		
		return $attachments;
	
	}

	/*
		Send email
		by David Guest
	*/
	
	function send_email($to='', $subject, $body, $cc='', $bcc='') {
		
		
		$this->setup();
		
		$CreateItem->MessageDisposition="SendAndSaveCopy";
		$CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = "sentitems";
		$CreateItem->Items->Message->Body = array("_" => $body, "BodyType" => "Text");
		$CreateItem->Items->Message->ItemClass = "IPM.Note";
		$CreateItem->Items->Message->Subject = $subject;
		$to = $this->parse_emails($to);
		$cc = $this->parse_emails($cc);
		$bcc = $this->parse_emails($bcc);
		for($t=0;$t<count($to);$t++) {
			$CreateItem->Items->Message->ToRecipients->Mailbox[$t]->EmailAddress = $to[$t];
		}
		for($c=0;$c<count($cc);$c++) {
			$CreateItem->Items->Message->CcRecipients->Mailbox[$c]->EmailAddress = $cc[$c];
		}
		for($b=0;$b<count($bcc);$b++) {
			$CreateItem->Items->Message->BccRecipients->Mailbox[$b]->EmailAddress = $bcc[$b];
		}
		$CreateItem->Items->Message->IsRead = false;
		$response = $this->client->CreateItem($CreateItem);

		$this->teardown();

		if($response->ResponseMessages->CreateItemResponseMessage->ResponseCode == "NoError") {
			return(array("outcome"=>"Message sent and a copy saved in Sent Items"));
		} else {
			return(array("outcome"=>"There was a problem sending the message"));
		}
		
	
	}
	
	/*
		Resolve name
		by David Guest
	*/
	
	function resolve_name($ambiguous) {
		$this->setup();
		
		$ResolveNames->ReturnFullContactData = true;
		$ResolveNames->UnresolvedEntry = $ambiguous;
		
		$response = $this->client->ResolveNames($ResolveNames);

		$this->teardown();

		return $response->ResponseMessages->ResolveNamesResponseMessage;
	
	}
	
	
	/*
	 	parse strings and email addresses
	 	internally used
	*/
	
	private function scrape($body,$start,$end) {
	
		$base = explode($start, $body);
		$core = explode($end, $base[1]);
		return $core[0];
		
	}
	
	// parse strings and return all instances
	private function scrape_all($body, $start, $end) {
	
		$values = array();
		if(stristr($body, $start)) {
			$base = explode($start, $body);
			for($i=1;$i<count($base);$i++) {
				if(stristr($base[$i], $end)) {
					$core = explode($end, $base[$i]);
					$values[] = $core[0];
				}
			}
		} 
		return $values;
	}

	//parse strings for email addresses
	private function parse_emails($str) {
	
		$str = str_replace('"', " ", $str);
		$str = str_replace("<", " ", $str); $str = str_replace(">", " ", $str);
		$str = str_replace("&lt;", " ", $str); $str = str_replace("&gt;", " ", $str);
		$str = str_replace("[", " ", $str); $str = str_replace("]", " ", $str);
		$str = str_replace("(", " ", $str); $str = str_replace(")", " ", $str);
		
		$pattern ="/(?:[a-zA-Z0-9!#$%&'*+=?^_`{|}~-]+(?:\.[a-zA-Z0-9!#$%&'*+=?^_`{|}~-]+)*|\"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*\")@(?:(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-zA-Z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])/";
	
		preg_match_all($pattern, $str, $matches);
		return $matches[0];
	
	}	

	/**
	 * Sets up stream handling. Internally used.
	 * 
	 * @access private
	 * @return void
	 */
	private function setup($specialised="") {

		
		$header = array();
		
		if($this->impersonate != NULL) {
			$impheader = new ImpersonationHeader($this->impersonate);
			$header[] = new SoapHeader("http://schemas.microsoft.com/exchange/services/2006/messages", "ExchangeImpersonation", $impheader, false);
			
		}

		
		//this header is required for some features
		//e.g. folder operations - added by Dave Guest
		$header[] = new SoapHeader("http://schemas.microsoft.com/exchange/services/2006/types", "RequestServerVersion", array("Version"=>$this->version));
		
		//add headers
		$this->client->__setSoapHeaders($header);
		
		
			
		stream_wrapper_unregister('http');
		stream_wrapper_register('http', 'ExchangeNTLMStream') or die("Failed to register protocol");

	}

	/**
	 * Tears down stream handling. Internally used.
	 * 
	 * @access private
	 * @return void
	 */
	function teardown() {
		stream_wrapper_restore('http');
	}
}

class ImpersonationHeader {

	var $ConnectingSID;

	function __construct($email) {
		$this->ConnectingSID->PrimarySmtpAddress = $email;
	}

}

class NTLMSoapClient extends SoapClient {
	function __doRequest($request, $location, $action, $version) {
		$headers = array(
			'Method: POST',
			'Connection: Keep-Alive',
			'User-Agent: PHP-SOAP-CURL',
			'Content-Type: text/xml; charset=utf-8',
			'SOAPAction: "'.$action.'"',
		);  
		$this->__last_request_headers = $headers;

		$ch = curl_init($location);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
		curl_setopt($ch, CURLOPT_POST, true );
		curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true );
		curl_setopt($ch, CURLOPT_POSTFIELDS, $request);
		curl_setopt($ch, CURLOPT_USERPWD, $this->user.':'.$this->password);
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
		curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
		curl_setopt($ch, CURLOPT_CONNECTTIMEOUT ,60); 
		//curl_setopt($ch, CURLOPT_VERBOSE, true);
		$response = curl_exec($ch);
		$c_error = curl_error($ch);
		
		
		
		
		return $response;
	}   
	function __getLastRequestHeaders() {
		return implode("n", $this->__last_request_headers)."n";
	}   
}

class ExchangeNTLMSoapClient extends NTLMSoapClient {
	public $user = '';
	public $password = '';
}

class NTLMStream {
	private $path;
	private $mode;
	private $options;
	private $opened_path;
	private $buffer;
	private $pos;

	public function stream_open($path, $mode, $options, $opened_path) {
		echo "[NTLMStream::stream_open] $path , mode=$mode n";
		$this->path = $path;
		$this->mode = $mode;
		$this->options = $options;
		$this->opened_path = $opened_path;
		$this->createBuffer($path);
		return true;
	}

	public function stream_close() {
		echo "[NTLMStream::stream_close] n";
		curl_close($this->ch);
	}

	public function stream_read($count) {
		echo "[NTLMStream::stream_read] $count n";
		if(strlen($this->buffer) == 0) {
			return false;
		}
		$read = substr($this->buffer,$this->pos, $count);
		$this->pos += $count;
		return $read;
	}

	public function stream_write($data) {
		echo "[NTLMStream::stream_write] n";
		if(strlen($this->buffer) == 0) {
			return false;
		}
		return true;
	}

	public function stream_eof() {
		echo "[NTLMStream::stream_eof] ";
		if($this->pos > strlen($this->buffer)) {
			echo "true n";
			return true;
		}
		echo "false n";
		return false;
	}

	/* return the position of the current read pointer */
	public function stream_tell() {
		echo "[NTLMStream::stream_tell] n";
		return $this->pos;
	}

	public function stream_flush() {
		echo "[NTLMStream::stream_flush] n";
		$this->buffer = null;
		$this->pos = null;
	}

	public function stream_stat() {
		echo "[NTLMStream::stream_stat] n";
		$this->createBuffer($this->path);
		$stat = array(
			'size' => strlen($this->buffer),
		);
		return $stat;
	}

	public function url_stat($path, $flags) {
		echo "[NTLMStream::url_stat] n";
		$this->createBuffer($path);
		$stat = array(
			'size' => strlen($this->buffer),
		);
		return $stat;
	}

	/* Create the buffer by requesting the url through cURL */
	private function createBuffer($path) {
		if($this->buffer) {
			return;
		}
		echo "[NTLMStream::createBuffer] create buffer from : $pathn";
		$this->ch = curl_init($path);
		curl_setopt($this->ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($this->ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
		curl_setopt($this->ch, CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
		curl_setopt($this->ch, CURLOPT_USERPWD, $this->user.':'.$this->password);
		echo $this->buffer = curl_exec($this->ch);
		echo "[NTLMStream::createBuffer] buffer size : ".strlen($this->buffer)."bytesn";
		$this->pos = 0;
	}
}

class ExchangeNTLMStream extends NTLMStream {
	protected $user = '';
	protected $password = '';
}


