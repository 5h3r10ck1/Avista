<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Adapter for Carousel RDA
 *
 * @author seth.phillips
 * 
 */


class RDA {
    
    function __construct($UserName, $Password, $ServerURI = 'localhost') {
        $this->UserName = $UserName;
        $this->Password = $Password;
        $this->HostIP = $ServerURI;
    }
    
    private $Port = 56906;
    
    public $HostIP = '';
    public $UserName = '';
    public $Password = '';
    
    private $ZoneID = array();
    private $ZoneTag = array();
    private $ZoneName = array();
    
    private $TemplateName = '';
    private $BulletinGuid = array();
    private $BulletinTag = array();
    private $CrawlGuid = array();
    private $Alert = false;
    
    private $Block = array();
    
    private $AlwaysOn = true;
    
    public $DateTimeOn = '';
    private $DateTimeOff = '';
    private $CycleTimeOn = '00:00:00';
    private $CycleTimeOff = '23:59:59';
    
    private $Weekdays = 127;
    
    public $WebEnabled = true;
    public $DisplayDuration = 30;
    
    private $LastError ='';
    private $CreatedBulletins = array();
    
    public $Description = 'RDA Created Bulletin';
    
    // Zone information Getters and Setters
    /**
     * Set Zone IDs
     * 
     * Allows the setting of 1 or more Zone IDs for use in the RDA.  
     * This function will overwrite the existing Zone IDs if any have been set. 
     * Related functions are getZoneIDs() and clearZoneIDs() 
     * 
     * @param int|string|array $ID a single Zone ID or an array of Zone IDs  
     * @return boolean returns false on failure.
     */
    public function setZoneIDs($ID = null){
        if($ID === null){return false;}
        $this->ZoneID = array();
        if(is_array($ID)){
            foreach($ID as $i){ array_push($this->ZoneID, $i); }
        }
        else{ array_push($this->ZoneID, $ID);}
    }
    
    /**
     * Get Zone IDs
     * 
     * Returns the list of Zone IDs 
     * Related functions are setZoneIDs() and clearZoneIDs() 
     *  
     * @return array array of Zone IDs
     */
    public function getZoneIDs(){
        return $this->ZoneID;
    }
    /**
     * Clear Zone IDs
     * 
     * Clears out one or all of the zone IDs that have been set by setZoneIDs().
     * 
     * @param string|int|null $ID A single zone ID or Null. Passing null will clear out all Zone IDs.
     * @return boolean returns true on success
     */
    public function clearZoneIDs($ID = null){
        if($ID === null){$this->ZoneID = array();return true;}
        else {foreach($this->ZoneID as $k=>$v){if($v===$ID){ array_splice($this->ZoneID, $k, 1);return TRUE;} } }
    }
    /**
     * Function to set Zone Tags
     * 
     * Allows the setting of 1 or more Zone Tags for use in the RDA.  
     * This function will overwrite the existing Zone Tags if any have been set. 
     * Related functions are getZoneTags() and clearZoneTags() 
     * 
     * @param int|string|array $Tag a single Zone tag or an array of Zone tags  
     * @return boolean returns false on failure.
     */
    public function setZoneTags($Tag = null){
        if($Tag === null){return false;}
        $this->ZoneTag = array();
        if(is_array($Tag)){
            foreach($Tag as $i){ array_push($this->ZoneTag, $i); }
        }
        else{ array_push($this->ZoneTag, $Tag);} 
    }
    /**
     * Get Zone Tags
     * 
     * Returns the list of Zone Tags 
     * Related functions are setZoneTags() and clearZoneTags() 
     *  
     * @return array array of Zone Tags
     */
    public function getZoneTags(){
        return $this->ZoneTag;
    }
    /**
     * Clear Zone Tags
     * 
     * Clears out one or all of the zone Tags that have been set by setZoneTags().
     * 
     * @param string|int|null $Tag A single zone Tag or Null. Passing null will clear out all Zone Tags.
     * @return boolean returns true on success
     */
    public function clearZoneTags($Tag = null){
        if($Tag === null){$this->ZoneTag = array();return true;}
        else {foreach($this->ZoneTag as $k=>$v){if($v===$Tag){ array_splice($this->ZoneTag, $k, 1);return TRUE;} } }
    }
    /**
     * Set Zone Names
     * 
     * Allows the setting of 1 or more Zone Names for use in the RDA.  
     * This function will overwrite the existing Zone Names if any have been set. 
     * Related functions are getZoneNames() and clearZoneNames() 
     * 
     * @param int|string|array $Name a single Zone Name or an array of Zone Names  
     * @return boolean returns false on failure.
     */
    public function setZoneNames($Name = null){
        if($Name === null){return false;}
        $this->ZoneName = array();
        if(is_array($Name)){
            foreach($Name as $i){ array_push($this->ZoneName, $i); }
        }
        else{ array_push($this->ZoneName, $Name);} 
    }
    /**
     * Get Zone Names
     * 
     * Returns the list of Zone Names 
     * Related functions are setZoneNames() and clearZoneNames() 
     *  
     * @return array array of Zone Names
     */
    public function getZoneNames(){
        return $this->ZoneName;
    }
    /**
     * Clear Zone Names
     * 
     * Clears out one or all of the zone Names that have been set by setZoneNames().
     * 
     * @param string|int|null $Name A single zone Name or Null. Passing null will clear out all Zone Names.
     * @return boolean returns true on success
     */
    public function clearZoneNames($Name = null){
        if($Name === null){$this->ZoneName = array();return true;}
        else {foreach($this->ZoneName as $k=>$v){if($v===$Name){ array_splice($this->ZoneName, $k, 1);return TRUE;} } }
    }
    /**
     * Reset Zone Information
     * 
     * Resets the list of Zone IDs Tags and Names to be used by the RDA.  The same as calling clearZoneIDs(), clearZoneTags() and clearZoneNames().
     * 
     */
    public function resetZones(){
        $this->ZoneID = array();
        $this->ZoneTag = array();
        $this->ZoneName = array();
    }
    // Template Getters and Setters
    /**
     * 
     * Set Template Name
     * 
     * Sets the template name to be used by CreatePage()
     * 
     * @param string $Name The name of the template to be used.
     * 
     */
    public function setTemplateName($Name){
        $this->TemplateName = $Name;
    }
    /**
     * 
     * Get Template Name
     * 
     * Gets the currently set template name to be used by CreatePage()
     * 
     * @return string Template Name
     * 
     */
    public function getTemplateName(){
        return $this->TemplateName;
    }
    
    /**
     * Set Bulletin Guids
     * 
     * Allows the setting of 1 or more Bulletin Guids for use in the RDA.  
     * This function will overwrite the existing Bulletin Guids if any have been set. 
     * Related functions are getBulletinGuids() and clearBulletinGuids() 
     * 
     * @param string|array $Guid a single Bulletin Guids or an array of Bulletin Guidss  
     * @return boolean returns false on failure.
     */
    public function setBulletinGuids($Guid = null){
        if($Guid === null){return false;}
        $this->BulletinGuid = array();
        if(is_array($Guid)){
            foreach($Guid as $i){ array_push($this->BulletinGuid, $i); }
        }
        else{ array_push($this->BulletinGuid, $Guid);}  
    }
    /**
     * Get Bulletin Guids
     * 
     * Returns the list of Bulletin Guids 
     * Related functions are setBulletinGuids() and clearBulletinGuids() 
     *  
     * @return array array of Bulletin Guids
     */
    public function getBulletinGuids(){
        return $this->BulletinGuid;
    }
    /**
     * Clear Bulletin Guids
     * 
     * Clears out one or all of the Bulletin Guids that have been set by setBulletinGuids().
     * 
     * @param string|null $Name A single Bulletin Guid or Null. Passing null will clear out all Bulletin Guids.
     * @return boolean returns true on success
     */
    public function clearBulletinGuids($Guid = null){
        if($Guid === null){$this->BulletinGuid = array();return true;}
        else {foreach($this->BulletinGuid as $k=>$v){if($v===$Guid){ array_splice($this->BulletinGuid, $k, 1);return TRUE;} } }
    }
    /**
     * Set Crawl Guids
     * 
     * Allows the setting of 1 or more Crawl Guids for use in the RDA.  
     * This function will overwrite the existing Crawl Guids if any have been set. 
     * Related functions are getCrawlGuids() and clearCrawlGuids() 
     * 
     * @param string|array $Guid a single Crawl Guid or an array of Crawl Guids  
     * @return boolean returns false on failure.
     */
    public function setCrawlGuids($Guid = null){
        if($Guid === null){return false;}
        $this->CrawlGuid = array();
        if(is_array($Guid)){
            foreach($Guid as $i){ array_push($this->CrawlGuid, $i); }
        }
        else{ array_push($this->CrawlGuid, $Guid);}  
    }
    /**
     * Get Crawl Guids
     * 
     * Returns the list of Crawl Guids 
     * Related functions are setCrawlGuids() and clearCrawlGuids() 
     *  
     * @return array array of Crawl Guids
     */
    public function getCrawlGuids(){
        return $this->CrawlGuid;
    }
    /**
     * Clear Crawl Guids
     * 
     * Clears out one or all of the Crawl Guids that have been set by setCrawlGuids().
     * 
     * @param string|null $Name A single Crawl Guid or Null. Passing null will clear out all Crawl Guids.
     * @return boolean returns true on success
     */
    public function clearCrawlGuids($Guid = null){
        if($Guid === null){$this->CrawlGuid = array();return true;}
        else {foreach($this->CrawlGuid as $k=>$v){if($v===$Guid){ array_splice($this->CrawlGuid, $k, 1);return TRUE;} } }
    }
    /**
     * Set Bulletin Tags
     * 
     * Allows the setting of 1 or more Bulletin Tags for use in the RDA.  
     * This function will overwrite the existing Bulletin Tags if any have been set. 
     * Related functions are getBulletinTags() and clearBulletinTags() 
     * 
     * @param string|array $Guid a single Bulletin Tag or an array of Bulletin Tags  
     * @return boolean returns false on failure.
     */
    public function setBulletinTags($Tag = null){
        if($Tag === null){return false;}
        $this->BulletinTag = array();
        if(is_array($Tag)){
            foreach($Tag as $i){ array_push($this->BulletinTag, $i); }
        }
        else{ array_push($this->BulletinTag, $Tag);}  
    }
    /**
     * Get Bulletin Tags
     * 
     * Returns the list of Bulletin Tags 
     * Related functions are setBulletinTags() and clearBulletinTags() 
     *  
     * @return array array of Bulletin Tags
     */
    public function getBulletinTags(){
        return $this->BulletinTag;
    }
    /**
     * Clear Bulletin Tags
     * 
     * Clears out one or all of the Bulletin Tags that have been set by setBulletinTags().
     * 
     * @param string|null $Name A single Bulletin Tag or Null. Passing null will clear out all Bulletin Tags.
     * @return boolean returns true on success
     */
    public function clearBulletinTags($Tag = null){
        if($Tag === null){$this->BulletinTag = array();return true;}
        else {foreach($this->BulletinTag as $k=>$v){if($v===$Tag){ array_splice($this->BulletinTag, $k, 1);return TRUE;} } }
    }
    /**
     * Set a Schedule
     * 
     * Sets a Schedule to be used when creating or updating a bulletin or crawl.
     * 
     * @param DateTime $DateOn DateTime Object representing when the schdule begins 
     * @param DateTime $DateOff DateTime Object representing when the schdule ends 
     * 
     */
    public function setSchedule(DateTime $DateOn, DateTime $DateOff){
        $this->AlwaysOn = false;
        $this->DateTimeOn = $DateOn->format(DateTime::W3C);
        $this->DateTimeOff = $DateOff->format(DateTime::W3C);
    }
    /**
     * Get the Schdule
     * 
     * Gets the current schedule start and end dates and times that will be used to create or update bulletins and crawls
     * 
     * @return string Representation of both the DateTimeOn and DateTimeOff separated by a en-dash, or 'Always On'. (ex.) '2014-04-05 00:00:00 - 2014-08-10 23:59:59' 
     * 
     */
    public function getSchedule(){
        if($this->AlwaysOn){return 'Always On';}
        return $this->DateTimeOn.' - '.$this->DateTimeOff;
    }
    /**
     * Clear Schedule
     * 
     * This defaults the RDA command to an Always On state.
     * 
     */
    public function clearSchedule(){
        $this->AlwaysOn = true;
    }
    /**
     * Set Cycle Time
     * 
     * Defines a start and end time to be used during each day of a bulletins schedule. Use a 24 hour time reference formatted 'HH:MM:SS' or 'HH:MM'.
     * 
     * @param string $TimeOn String formatted as 'HH:MM:SS' 
     * @param string $TimeOff String formatted as 'HH:MM:SS'
     * 
     * @return boolean returns true on success or false on failure. Failure will not reset the Cycle Times.
     */
    public function setCycleTime($TimeOn = '00:00:00',$TimeOff = '23:59:59'){
        if(preg_match('/^[0-9]{1,2}:[0-9]{2}$/',$TimeOn) === 1){$TimeOn.=':00';}
        if(preg_match('/^[0-9]{1,2}:[0-9]{2}$/',$TimeOff) === 1){$TimeOff.=':00';}
        
        if(preg_match('/^[0-9]:[0-9]{2}:[0-9]{2}$/',$TimeOn) === 1){$TimeOn ='0'.$TimeOn;}
        if(preg_match('/^[0-9]:[0-9]{2}:[0-9]{2}$/',$TimeOff) === 1){$TimeOff ='0'.$TimeOff;}
        
        if(preg_match('/^[0-9]{1,2}:[0-9]{2}:[0-9]{2}$/',$TimeOff) === 1 && preg_match('/^[0-9]{1,2}:[0-9]{2}:[0-9]{2}$/',$TimeOn) === 1){
            $this->CycleTimeOn = $TimeOn;
            $this->CycleTimeOff = $TimeOff;
            return 1;    
        }
        else return 0;
    }
    /**
     * 
     * Get Cycle Time
     * 
     * Gets the currently set cycle time that will be used to create or update bulletins and crawls.
     * 
     * @return string Cycle Time On and Off separated by an en-dash. (ex.) "08:30:00 - 23:15:00"
     * 
     */
    public function getCycleTime(){
        return $this->CycleTimeOn.' - '.$this->CycleTimeOff;
    }
    /**
     * Clear Schedule
     * 
     * This defaults the Cycle Time to an Always On state by setting CycleTimeOn and CycleTimeOff to '00:00:00' and '23:59:59' respectively.
     * 
     */
    public function clearCycleTime(){
        $this->CycleTimeOn = '00:00:00';
        $this->CycleTimeOff = '23:59:59';
    }
    /**
     * Set Week Days
     * 
     * Sets the valid playback days for use within a bulletins schedule.  Most abbreviations or full day names are valid.  Days should be separated by a ' ', '-', ',' or ':'.  Do not mix separators.
     * 
     * @param string $Days A properly formed string representation of the Days of the week. (ex.) 'M-T-W-Th-F' or 'Mon,Tuesday,Wed,Thur,Friday' or a blank string for none
     * 
     * @return boolean Returns true if weekdays have been set.
     * Returns false if no days could be parsed and weekdays have been cleared.  
     */
    public function setWeekDays($Days = 'Sat-Mon-Tue-Wed-Thur-Fri-Sat'){
        if(preg_match('/:/', $Days) !== 0){$X = explode(':', $Days);}
        elseif(preg_match('/-/', $Days) !== 0){$X = explode('-', $Days);}
        elseif(preg_match("/,/", $Days) !== 0){$X = explode(',', $Days);}
        elseif(preg_match('/ /', $Days) !== 0){$X = explode(' ', $Days);}
        else{$X = array($Days);}
        $Y = array(false,false,false,false,false,false,false);
        $result = '';
        
        foreach($X as $x){
            switch(trim($x)){
                case 'Sat':
                    $Y[0] = true;
                    break;
                case 'Saturday':
                    $Y[0] = true;
                    break;
                case 'Fri':
                    $Y[1] = true;
                    break;
                case 'Friday':
                    $Y[1] = true;
                    break;
                case 'F':
                    $Y[1] = true;
                    break;
                case 'Thur':
                    $Y[2] = true;
                    break;
                case 'Th':
                    $Y[2] = true;
                    break;
                case 'Thurs':
                    $Y[2] = true;
                    break;
                case 'Thursday':
                    $Y[2] = true;
                    break;
                case 'Wed':
                    $Y[3] = true;
                    break;
                case 'Wednesday':
                    $Y[3] = true;
                    break;
                case 'W':
                    $Y[3] = true;
                    break;
                case 'Tue':
                    $Y[4] = true;
                    break;
                case 'Tuesday':
                    $Y[4] = true;
                    break;
                case 'Tues':
                    $Y[4] = true;
                    break;
                case 'T':
                    $Y[4] = true;
                    break;
                case 'Mon':
                    $Y[5] = true;
                    break;
                case 'Monday':
                    $Y[5] = true;
                    break;
                case 'M':
                    $Y[5] = true;
                    break;
                case 'Sun':
                    $Y[6] = true;
                    break;
                case 'Sunday':
                    $Y[6] = true;
                    break;
                default:
                    break;
            }
        }
        foreach($Y as $y){
            $result .= ($y)? '1':'0';
        }
        $DayBits = base_convert($result,2,10);
        
        $this->Weekdays = $DayBits;
        if($DayBits === 0){
            return 0;
        }
        else {
            return 1;
        }
    }
    
    /**
     * Set Block
     * 
     * Prepares a Template block for use in Creating or Updating Bulletins.  Blocks have a name and a value.  Name is the block name as set in the Carousel Template Editor
     * and value is the text in a text block, the uri in an RSS or web picture block and the GUID of a video or picture block.
     * 
     * @param string $name the name of the template or bulletin block
     * @param string $value the value of the text block, a uri for web image and rss blocks or a GUID for picture and video blocks
     */
    public function setBlock($name, $value){
        $block['Name'] = $name;
        $block['Value'] = $value;
        array_push($this->Block,$block);
    }
    /**
     * 
     * Get Blocks
     * 
     * Gets and array of the currently set blocks.
     * 
     * @return array Array of the currently set blocks 
     * 
     */
    public function getBlocks(){
        return $this->Block;
    }
    /**
     * 
     * @param int|null $key integer representing the key of the block as referenced in getBlocks(), or null will clear out all blocks.
     * @return boolean returns true for no good reason.
     */
    public function clearBlocks($key = null){
        if($key === null){$this->Block = array();return true;}
        array_splice($this->Block, $key, 1);
        return true;
    }
    /**
     * Is Alert
     * 
     * Gets or Sets the Alerting Feature. If a Boolean is supplied as an argument, it will set the current alert status
     * if no argument is supplied it will tell you if the current RDA is set to prepare an alert bulletin.
     * 
     * @param boolean|null $boolean Sets the current alert status if supplied.
     * @return boolean Returns the current value of the alert setting if no arguement is passed in
     */
    public function IsAlert($boolean = ''){
        if(isset($boolean) && is_bool($boolean)){
            $this->Alert = $boolean;
        }
        else{
            return $this->Alert;
        }
    }
    
    /**
     * Create Page
     * 
     * Creates Bulletins in Carousel from a template.  This requires zone information, a template name, and 
     * and matching blocks to be set.  Optional scheduling information can also be set.  The command takes no arguments and requires
     * that all needed information be set in advance using other methods. One bulletin will be created for each zone and each zone supplied
     * must have the template available in it.
     * 
     * @return array|boolean reutrns an array of created bulletin GUIDs on success, or false on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function CreatePage(){
        $command = 'CreatePage';
        $xml = new DOMDocument('1.0','utf-8');
        $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
        $xml->appendChild($root);
        
        $ActionRoot = $xml->createElement($command);
        $root->appendChild($ActionRoot);
        
        $UserName = $xml->createElement('UserName',$this->UserName);
        $ActionRoot->appendChild($UserName);
        $Password = $xml->createElement('Password',$this->Password);
        $ActionRoot->appendChild($Password);
        
        $ZoneSet = $xml->createElement('ZoneSet');
        foreach($this->ZoneID as $x){
            $ZoneID = $xml->createElement('ZoneID',$x);
            $ZoneSet->appendChild($ZoneID);
        }
        foreach($this->ZoneName as $x){
            $ZoneName = $xml->createElement('ZoneName');
            $ZoneNameText = $xml->createTextNode($x);
            $ZoneName->appendChild($ZoneNameText);
            $ZoneSet->appendChild($ZoneName);
        }
        foreach($this->ZoneTag as $x){
            $ZoneTag = $xml->createElement('ZoneTag',$x);
            $ZoneSet->appendChild($ZoneTag);
        }
        $ActionRoot->appendChild($ZoneSet);
        
        $_AlwaysOn = ($this->AlwaysOn)?'true':'false';
        $AlwaysOn = $xml->createElement('AlwaysOn', $_AlwaysOn);
        $ActionRoot->appendChild($AlwaysOn);
        if(!$this->AlwaysOn){
            $DateTimeOn = $xml->createElement('DateTimeOn', $this->DateTimeOn);
            $DateTimeOff = $xml->createElement('DateTimeOff', $this->DateTimeOff);
            $ActionRoot->appendChild($DateTimeOn);
            $ActionRoot->appendChild($DateTimeOff);
        }
        $CycleTimeOn = $xml->createElement('CycleTimeOn',$this->CycleTimeOn);
        $ActionRoot->appendChild($CycleTimeOn);
        $CycleTimeOff = $xml->createElement('CycleTimeOff',$this->CycleTimeOff);
        $ActionRoot->appendChild($CycleTimeOff);
        $Weekdays = $xml->createElement('Weekdays',$this->Weekdays);
        $ActionRoot->appendChild($Weekdays);
        
        //$DisplayDuration = $xml->createElement('DisplayDuration',$this->DisplayDuration);
        //$ActionRoot->appendChild($DisplayDuration);
        
        $_WebEnabled = ($this->WebEnabled)?'true':'false';
        $WebEnabled = $xml->createElement('WebEnabled',$_WebEnabled);
        $ActionRoot->appendChild($WebEnabled);
        
        $Description = $xml->createElement('Description');
        $DescriptionText = $xml->createTextNode($this->Description);
        $Description->appendChild($DescriptionText);
        $ActionRoot->appendChild($Description);
        
        if($this->Alert)
            $PT = 'Alert';
        else
            $PT = 'Standard';
        $PageType = $xml->createElement('PageType',$PT);
        $ActionRoot->appendChild($PageType);
        
        $PageTemplate = $xml->createElement('PageTemplate');
        $TemplateName = $xml->createElement('TemplateName',$this->TemplateName);
        $PageTemplate->appendChild($TemplateName);
        foreach($this->Block as $block){
            $BlockNode = $xml->createElement('Block');
            $Name = $xml->createAttribute('Name');
            $Value = $xml->createAttribute('Value');
            $Name->value = $block['Name'];
            $Value->value = $block['Value'];
            $BlockNode->appendChild($Name);
            $BlockNode->appendChild($Value);
            $PageTemplate->appendChild($BlockNode);
        }
        
        $ActionRoot->appendChild($PageTemplate);
        
        return $this->fireRDA($xml,$command);
    }
    /**
     * Update Page
     * 
     * Update a Page or 'Bulletin' in Carousel from a template.  This requires a bulletin GUID, or 
     * bulletin tags as well as matching blocks to be set.  Optional scheduling information can also be set.  The command takes no arguments and requires
     * that all needed information be set in advance using other methods.
     * 
     * @return array|boolean reutrns an array of effected bulletin GUIDs on success, or false on failure.
     * Upon failure an error message will be available in getLastError(); NOTE: if supplied an invalid GUID, the system
     * will behave as if a bulletin was updated.
     * 
     */
    public function UpdatePage(){
        $command = 'UpdatePage';
        $BulletinList = array();
        
        foreach($this->BulletinGuid as $GUID){
            $xml = new DOMDocument('1.0','utf-8');
            $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
            $xml->appendChild($root);

            $ActionRoot = $xml->createElement($command);
            $root->appendChild($ActionRoot);

            $UserName = $xml->createElement('UserName',$this->UserName);
            $ActionRoot->appendChild($UserName);
            $Password = $xml->createElement('Password',$this->Password);
            $ActionRoot->appendChild($Password);

            $UpdateGUID = $xml->createElement('UpdateGUID',$GUID);
            $ActionRoot->appendChild($UpdateGUID);

            $_AlwaysOn = ($this->AlwaysOn)?'true':'false';
            $AlwaysOn = $xml->createElement('AlwaysOn', $_AlwaysOn);
            $ActionRoot->appendChild($AlwaysOn);
            if(!$this->AlwaysOn){
                $DateTimeOn = $xml->createElement('DateTimeOn', $this->DateTimeOn);
                $DateTimeOff = $xml->createElement('DateTimeOff', $this->DateTimeOff);
                $ActionRoot->appendChild($DateTimeOn);
                $ActionRoot->appendChild($DateTimeOff);
            }
            $CycleTimeOn = $xml->createElement('CycleTimeOn',$this->CycleTimeOn);
            $ActionRoot->appendChild($CycleTimeOn);
            $CycleTimeOff = $xml->createElement('CycleTimeOff',$this->CycleTimeOff);
            $ActionRoot->appendChild($CycleTimeOff);
            $Weekdays = $xml->createElement('Weekdays',$this->Weekdays);
            $ActionRoot->appendChild($Weekdays);

            //$DisplayDuration = $xml->createElement('DisplayDuration',$this->DisplayDuration);
            //$ActionRoot->appendChild($DisplayDuration);

            $_WebEnabled = ($this->WebEnabled)?'true':'false';
            $WebEnabled = $xml->createElement('WebEnabled',$_WebEnabled);
            $ActionRoot->appendChild($WebEnabled);

            $Description = $xml->createElement('Description');
            $DescriptionText = $xml->createTextNode($this->Description);
            $Description->appendChild($DescriptionText);
            $ActionRoot->appendChild($Description);

            if($this->Alert)
                $PT = 'Alert';
            else
                $PT = 'Standard';
            $PageType = $xml->createElement('PageType',$PT);
            $ActionRoot->appendChild($PageType);

            foreach($this->Block as $block){
                $BlockNode = $xml->createElement('Block');
                $Name = $xml->createAttribute('Name');
                $Value = $xml->createAttribute('Value');
                $Name->value = $block['Name'];
                $Value->value = $block['Value'];
                $BlockNode->appendChild($Name);
                $BlockNode->appendChild($Value);
                $ActionRoot->appendChild($BlockNode);
            }

            $result = $this->fireRDA($xml,$command);
            if($result){
                foreach($result as $x){array_push($BulletinList,$x);} 
            }
        }
        
        if( count($this->BulletinTag) > 0 ){
            $xml = new DOMDocument('1.0','utf-8');
            $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
            $xml->appendChild($root);

            $ActionRoot = $xml->createElement($command);
            $root->appendChild($ActionRoot);

            $UserName = $xml->createElement('UserName',$this->UserName);
            $ActionRoot->appendChild($UserName);
            $Password = $xml->createElement('Password',$this->Password);
            $ActionRoot->appendChild($Password);

            $SelectBulletinTags = $xml->createElement('SelectBulletinTags');
                foreach($this->BulletinTag as $tag){
                    $Tag = $xml->createElement('Tag');
                    $TagText = $xml->createTextNode($tag);
                    $Tag->appendChild($TagText);
                    $SelectBulletinTags->appendChild($Tag);
                }
            $ActionRoot->appendChild($SelectBulletinTags);

            $_AlwaysOn = ($this->AlwaysOn)?'true':'false';
            $AlwaysOn = $xml->createElement('AlwaysOn', $_AlwaysOn);
            $ActionRoot->appendChild($AlwaysOn);
            if(!$this->AlwaysOn){
                $DateTimeOn = $xml->createElement('DateTimeOn', $this->DateTimeOn);
                $DateTimeOff = $xml->createElement('DateTimeOff', $this->DateTimeOff);
                $ActionRoot->appendChild($DateTimeOn);
                $ActionRoot->appendChild($DateTimeOff);
            }
            $CycleTimeOn = $xml->createElement('CycleTimeOn',$this->CycleTimeOn);
            $ActionRoot->appendChild($CycleTimeOn);
            $CycleTimeOff = $xml->createElement('CycleTimeOff',$this->CycleTimeOff);
            $ActionRoot->appendChild($CycleTimeOff);
            $Weekdays = $xml->createElement('Weekdays',$this->Weekdays);
            $ActionRoot->appendChild($Weekdays);

            //$DisplayDuration = $xml->createElement('DisplayDuration',$this->DisplayDuration);
            //$ActionRoot->appendChild($DisplayDuration);

            $_WebEnabled = ($this->WebEnabled)?'true':'false';
            $WebEnabled = $xml->createElement('WebEnabled',$_WebEnabled);
            $ActionRoot->appendChild($WebEnabled);

            $Description = $xml->createElement('Description');
            $DescriptionText = $xml->createTextNode($this->Description);
            $Description->appendChild($DescriptionText);
            $ActionRoot->appendChild($Description);

            if($this->Alert)
                $PT = 'Alert';
            else
                $PT = 'Standard';
            $PageType = $xml->createElement('PageType',$PT);
            $ActionRoot->appendChild($PageType);

            foreach($this->Block as $block){
                $BlockNode = $xml->createElement('Block');
                $Name = $xml->createAttribute('Name');
                $Value = $xml->createAttribute('Value');
                $Name->value = $block['Name'];
                $Value->value = $block['Value'];
                $BlockNode->appendChild($Name);
                $BlockNode->appendChild($Value);
                $ActionRoot->appendChild($BlockNode);
            }

            $result = $this->fireRDA($xml,$command);
            if($result){
                foreach($result as $x){array_push($BulletinList,$x);} 
            }
        }
        if( count($BulletinList) === 0 ){return 0;}
        else {return $BulletinList;}
        
    }
    /**
     * Change Page Status
     * 
     * Changes the "status" of a bulletin.  This requires bulletin GUIDs or Bulletin Tags to be set and a status of "on" or "off" supplied as an argument.
     * 
     * 
     * @return array reutrns an associative array of GUID->true on success and GUID->false on failure. Tags will come back under TagList->boolean
     * Upon failure an error message will be available in getLastError();
     */
    public function ChangePageStatus($Status){ // 'on' or 'off'
        $command = 'ChangePageStatus';
        $responseList = array();
        foreach($this->BulletinGuid as $GUID){
            $xml = new DOMDocument('1.0','utf-8');
            $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
            $xml->appendChild($root);

            $ActionRoot = $xml->createElement($command);
            $root->appendChild($ActionRoot);

            $UserName = $xml->createElement('UserName',$this->UserName);
            $ActionRoot->appendChild($UserName);
            $Password = $xml->createElement('Password',$this->Password);
            $ActionRoot->appendChild($Password);

            $PageGUID = $xml->createElement('GUID',$GUID);
            $ActionRoot->appendChild($PageGUID);
            
            $pageStatus = $xml->createElement('Status',$Status);
            $ActionRoot->appendChild($pageStatus);
            
            $response = $this->fireRDA($xml,$command);
            $responseList[$GUID] = $response;
        }
        
        if( count($this->BulletinTag) > 0 ){
            $xml = new DOMDocument('1.0','utf-8');
            $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
            $xml->appendChild($root);

            $ActionRoot = $xml->createElement($command);
            $root->appendChild($ActionRoot);

            $UserName = $xml->createElement('UserName',$this->UserName);
            $ActionRoot->appendChild($UserName);
            $Password = $xml->createElement('Password',$this->Password);
            $ActionRoot->appendChild($Password);
            
            $SelectBulletinTags = $xml->createElement('SelectBulletinTags');
            foreach($this->BulletinTag as $tag){
                $Tag = $xml->createElement('Tag');
                $TagText = $xml->createTextNode($tag);
                $Tag->appendChild($TagText);
                $SelectBulletinTags->appendChild($Tag);
            }
            $ActionRoot->appendChild($SelectBulletinTags);
            
            $pageStatus = $xml->createElement('Status',$Status);
            $ActionRoot->appendChild($pageStatus);
       
            $response = $this->fireRDA($xml,$command);
            $responseList['TagList'] = $response;
        }
        if(count($responseList)>0){return $responseList;}
        else{return 0;}
    }
    /**
     * Change Page Status
     * 
     * Changes the "status" of a bulletin.  This requires bulletin GUIDs or Bulletin Tags to be set and an integer representation of status supplied as an argument.
     * Current=1, Queued=2, Hold=4, Saved=8, Stale=16
     * 
     * @param int $Status Integer representation of bulletin status Current=1, Queued=2, Hold=4, Saved=8, Stale=16
     * NOTE: it is not possible to set a bulletin that would otherwise be current to queued status or vice versa.  Instead used Stale or Saved or change schedule information.
     * 
     * @return array reutrns an associative array of GUID->true on success and GUID->false on failure. Tags will come back under TagList->boolean
     * Upon failure an error message will be available in getLastError();
     */
    public function SetPageStatus($Status){//Current=1, Queued=2, Hold=4, Saved=8, Stale=16
        $command = 'SetPageStatus';
        $responseList = array();
        foreach($this->BulletinGuid as $GUID){
            $xml = new DOMDocument('1.0','utf-8');
            $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
            $xml->appendChild($root);

            $ActionRoot = $xml->createElement($command);
            $root->appendChild($ActionRoot);

            $UserName = $xml->createElement('UserName',$this->UserName);
            $ActionRoot->appendChild($UserName);
            $Password = $xml->createElement('Password',$this->Password);
            $ActionRoot->appendChild($Password);

            $PageGUID = $xml->createElement('GUID',$GUID);
            $ActionRoot->appendChild($PageGUID);
            
            $pageStatus = $xml->createElement('Status',$Status);
            $ActionRoot->appendChild($pageStatus);
            
            $response = $this->fireRDA($xml,$command);
            $responseList[$GUID] = $response;
        }
        if(count($responseList)>0){return $responseList;}
        else{return 0;}
    }
    /**
     * Deactivate All Alert Pages
     * 
     * Deactivates all active alerts in a set of Zones.  Requires that zone information be set.
     * 
     * @return boolean returns true on success and false on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function DeactivateAllAlertPages(){
       $command = 'DeactivateAllAlertPages';
       $xml = new DOMDocument('1.0','utf-8');
       $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
       $xml->appendChild($root);
        
       $ActionRoot = $xml->createElement($command);
       $root->appendChild($ActionRoot);
       
       $UserName = $xml->createElement('UserName',$this->UserName);
       $ActionRoot->appendChild($UserName);
       $Password = $xml->createElement('Password',$this->Password);
       $ActionRoot->appendChild($Password);
        
       $ZoneSet = $xml->createElement('ZoneSet');
        foreach($this->ZoneID as $x){
            $ZoneID = $xml->createElement('ZoneID',$x);
            $ZoneSet->appendChild($ZoneID);
        }
        foreach($this->ZoneName as $x){
            $ZoneName = $xml->createElement('ZoneName');
            $ZoneNameText = $xml->createTextNode($x);
            $ZoneName->appendChild($ZoneNameText);
            $ZoneSet->appendChild($ZoneName);
        }
        foreach($this->ZoneTag as $x){
            $ZoneTag = $xml->createElement('ZoneTag',$x);
            $ZoneSet->appendChild($ZoneTag);
        }
        $ActionRoot->appendChild($ZoneSet);
    
        return $this->fireRDA($xml);
        
    }
    /**
     * Archive Page
     * 
     * Moves a Bulletin into Stale Bulletins.  Requires a bulletin GUID to be supplied as an argument.
     * 
     * @param string $GUID Guid of bulletin to be archived
     * @return boolean returns true on success or false on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function ArchivePage($GUID){
        $command = 'ArchivePage';
        $xml = new DOMDocument('1.0','utf-8');
        $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
        $xml->appendChild($root);
        
        $ActionRoot = $xml->createElement($command);
        $root->appendChild($ActionRoot);
        
        $UserName = $xml->createElement('UserName',$this->UserName);
        $ActionRoot->appendChild($UserName);
        $Password = $xml->createElement('Password',$this->Password);
        $ActionRoot->appendChild($Password);
        
        $PageGUID = $xml->createElement('GUID',$GUID);
        $ActionRoot->appendChild($PageGUID);
        
        return $this->fireRDA($xml,$command);
        
    }
    /**
     * Delete Page
     * 
     * Deletes a Bulletin.  Requires a bulletin GUID to be supplied as an argument.
     * 
     * @param string $GUID Guid of bulletin to be deleted
     * @return boolean returns true on success or false on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function DeletePage($GUID){
        $command = 'DeletePage';
        $xml = new DOMDocument('1.0','utf-8');
        $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
        $xml->appendChild($root);
        
        $ActionRoot = $xml->createElement($command);
        $root->appendChild($ActionRoot);
        
        $UserName = $xml->createElement('UserName',$this->UserName);
        $ActionRoot->appendChild($UserName);
        $Password = $xml->createElement('Password',$this->Password);
        $ActionRoot->appendChild($Password);
        
        $PageGUID = $xml->createElement('GUID',$GUID);
        $ActionRoot->appendChild($PageGUID);
        
        return $this->fireRDA($xml,$command);
       
    }
    /**
     * 
     * Delete All User Pages
     * 
     * Deletes ALL Pages created by the current user as defined in the RDA object creation.  
     * NOTE: BE CAREFULL as there is no confirmation and if logged in as an admin user or non RDA specific account,
     * you are liable to delete all bulletins on the system.  THIS IS NOT UNDOABLE.
     * 
     * @return boolean returns true on success or false on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function DeleteAllUserPages(){
        $command = 'DeleteAllUserPages';
        $xml = new DOMDocument('1.0','utf-8');
        $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
        $xml->appendChild($root);
        
        $ActionRoot = $xml->createElement($command);
        $root->appendChild($ActionRoot);
        
        $UserName = $xml->createElement('UserName',$this->UserName);
        $ActionRoot->appendChild($UserName);
        $Password = $xml->createElement('Password',$this->Password);
        $ActionRoot->appendChild($Password);
        
        return $this->fireRDA($xml,$command);
    }
    /**
     * 
     * Update Crawl
     * 
     * Updates the text in a crawl bulletin.  This requires that a GUID representing the crawl be set.  As of Carousel 6.4.5, there is a bug in Carousl that prevents this update from
     * being sent to the player.  The Server will update crawl text but the player will not.  It is best to Delete/Create new.
     * 
     * @param string Crawl Text
     * 
     * @return boolean|array returns an array of GUIDS on success or false on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function UpdateCrawl($Text = ''){
        $command = 'CreateCrawl';
        $CrawlList = array();
        foreach($this->CrawlGuid as $GUID){
            $xml = new DOMDocument('1.0','utf-8');
            $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
            $xml->appendChild($root);

            $ActionRoot = $xml->createElement($command);
            $root->appendChild($ActionRoot);

            $UserName = $xml->createElement('UserName',$this->UserName);
            $ActionRoot->appendChild($UserName);
            $Password = $xml->createElement('Password',$this->Password);
            $ActionRoot->appendChild($Password);

            $UpdateGUID = $xml->createElement('UpdateGUID',$GUID);
            $ActionRoot->appendChild($UpdateGUID);

            $CrawlText = $xml->createElement('CrawlText');
            $CrawlTextText = $xml->createTextNode($Text);
            $CrawlText->appendChild($CrawlTextText);
            $ActionRoot->appendChild($CrawlText);

            $ZoneSet = $xml->createElement('ZoneSet');
            foreach($this->ZoneID as $x){
                $ZoneID = $xml->createElement('ZoneID',$x);
                $ZoneSet->appendChild($ZoneID);
            }
            foreach($this->ZoneName as $x){
                $ZoneName = $xml->createElement('ZoneName');
                $ZoneNameText = $xml->createTextNode($x);
                $ZoneName->appendChild($ZoneNameText);
                $ZoneSet->appendChild($ZoneName);
            }
            foreach($this->ZoneTag as $x){
                $ZoneTag = $xml->createElement('ZoneTag',$x);
                $ZoneSet->appendChild($ZoneTag);
            }
            $ActionRoot->appendChild($ZoneSet);
            
            $_AlwaysOn = ($this->AlwaysOn)?'true':'false';
            $AlwaysOn = $xml->createElement('AlwaysOn', $_AlwaysOn);
            $ActionRoot->appendChild($AlwaysOn);
            if(!$this->AlwaysOn){
                $DateTimeOn = $xml->createElement('DateTimeOn', $this->DateTimeOn);
                $DateTimeOff = $xml->createElement('DateTimeOff', $this->DateTimeOff);
                $ActionRoot->appendChild($DateTimeOn);
                $ActionRoot->appendChild($DateTimeOff);
            }
            $CycleTimeOn = $xml->createElement('CycleTimeOn',$this->CycleTimeOn);
            $ActionRoot->appendChild($CycleTimeOn);
            $CycleTimeOff = $xml->createElement('CycleTimeOff',$this->CycleTimeOff);
            $ActionRoot->appendChild($CycleTimeOff);
            $Weekdays = $xml->createElement('Weekdays',$this->Weekdays);
            $ActionRoot->appendChild($Weekdays);

            //$DisplayDuration = $xml->createElement('DisplayDuration',$this->DisplayDuration);
            //$ActionRoot->appendChild($DisplayDuration);

            $_WebEnabled = ($this->WebEnabled)?'true':'false';
            $WebEnabled = $xml->createElement('WebEnabled',$_WebEnabled);
            $ActionRoot->appendChild($WebEnabled);

            $Description = $xml->createElement('Description');
            $DescriptionText = $xml->createTextNode($this->Description);
            $Description->appendChild($DescriptionText);
            $ActionRoot->appendChild($Description);

            $result = $this->fireRDA($xml,$command);
            if($result){
                foreach($result as $x){array_push($CrawlList,$x);}
            }
        }
        if( count($CrawlList) === 0 ){return 0;}
        else {return $CrawlList;}
    }
    /**
     * 
     * Create Crawl
     * 
     * Creates a crawl bulletin.  Requires that at least 1 Crawl Zone Tag, ID or Name be set.
     * 
     * @param string Crawl Text
     * 
     * @return boolean|array returns an array of GUIDS on success or false on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function CreateCrawl($Text = ''){
        $command = 'CreateCrawl';
        $xml = new DOMDocument('1.0','utf-8');
        $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
        $xml->appendChild($root);
        
        $ActionRoot = $xml->createElement($command);
        $root->appendChild($ActionRoot);
        
        $UserName = $xml->createElement('UserName',$this->UserName);
        $ActionRoot->appendChild($UserName);
        $Password = $xml->createElement('Password',$this->Password);
        $ActionRoot->appendChild($Password);
        
        $CrawlText = $xml->createElement('CrawlText');
        $CrawlTextText = $xml->createTextNode($Text);
        $CrawlText->appendChild($CrawlTextText);
        $ActionRoot->appendChild($CrawlText);
        
        $ZoneSet = $xml->createElement('ZoneSet');
        foreach($this->ZoneID as $x){
            $ZoneID = $xml->createElement('ZoneID',$x);
            $ZoneSet->appendChild($ZoneID);
        }
        foreach($this->ZoneName as $x){
            $ZoneName = $xml->createElement('ZoneName');
            $ZoneNameText = $xml->createTextNode($x);
            $ZoneName->appendChild($ZoneNameText);
            $ZoneSet->appendChild($ZoneName);
        }
        foreach($this->ZoneTag as $x){
            $ZoneTag = $xml->createElement('ZoneTag',$x);
            $ZoneSet->appendChild($ZoneTag);
        }
        $ActionRoot->appendChild($ZoneSet);
        
        $_AlwaysOn = ($this->AlwaysOn)?'true':'false';
        $AlwaysOn = $xml->createElement('AlwaysOn', $_AlwaysOn);
        $ActionRoot->appendChild($AlwaysOn);
        if(!$this->AlwaysOn){
            $DateTimeOn = $xml->createElement('DateTimeOn', $this->DateTimeOn);
            $DateTimeOff = $xml->createElement('DateTimeOff', $this->DateTimeOff);
            $ActionRoot->appendChild($DateTimeOn);
            $ActionRoot->appendChild($DateTimeOff);
        }
        $CycleTimeOn = $xml->createElement('CycleTimeOn',$this->CycleTimeOn);
        $ActionRoot->appendChild($CycleTimeOn);
        $CycleTimeOff = $xml->createElement('CycleTimeOff',$this->CycleTimeOff);
        $ActionRoot->appendChild($CycleTimeOff);
        $Weekdays = $xml->createElement('Weekdays',$this->Weekdays);
        $ActionRoot->appendChild($Weekdays);
        
        //$DisplayDuration = $xml->createElement('DisplayDuration',$this->DisplayDuration);
        //$ActionRoot->appendChild($DisplayDuration);
        
        $_WebEnabled = ($this->WebEnabled)?'true':'false';
        $WebEnabled = $xml->createElement('WebEnabled',$_WebEnabled);
        $ActionRoot->appendChild($WebEnabled);
        
        $Description = $xml->createElement('Description');
        $DescriptionText = $xml->createTextNode($this->Description);
        $Description->appendChild($DescriptionText);
        $ActionRoot->appendChild($Description);
        
        return $this->fireRDA($xml,$command);
    }
    /**
     * Get Zone List
     * 
     * Gets the zone information for the entire server.
     * 
     * @return boolean|array Array of Zone information associative arrays on success or false on failure
     * Zone information array contains 'ZoneID', 'ZoneName', and 'ZoneType'.
     * Upon failure an error message will be available in getLastError();
     */
    public function GetZoneList(){
       $command =  'GetZoneList';
       
       $xml = new DOMDocument('1.0','utf-8');
       $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
       $xml->appendChild($root);

       $ActionRoot = $xml->createElement($command);
       $root->appendChild($ActionRoot);
       $UserName = $xml->createElement('UserName',$this->UserName);
       $ActionRoot->appendChild($UserName);
       $Password = $xml->createElement('Password',$this->Password);
       $ActionRoot->appendChild($Password);
       
       return $this->fireRDA($xml,$command);
       
    }
    /**
     * Get Template List
     * 
     * Gets the Template information for select zone IDs.  Requires that Zone IDs be set.
     * 
     * @return boolean|array Returns an associative array of Template information organized by Zone ID
     * Template information consists of 'TemplateName' and a 'Block' array with each value having 'Name' 'Value' and 'Type' attributes
     * Returns False on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function GetTemplateList(){
        $command =  'GetTemplateList';
       $templateList = array();
       if(count($this->ZoneID) === 0){
           $this->LastError = "GetTemplateList() requires that at least one ZoneID be set";
           return 0;
       }
        foreach($this->ZoneID as $ID){
            $xml = new DOMDocument('1.0','utf-8');
           $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
           $xml->appendChild($root);

           $ActionRoot = $xml->createElement($command);
           $root->appendChild($ActionRoot);
           $UserName = $xml->createElement('UserName',$this->UserName);
           $ActionRoot->appendChild($UserName);
           $Password = $xml->createElement('Password',$this->Password);
           $ActionRoot->appendChild($Password);

           $ZoneID = $xml->createElement('ZoneID',$ID);
           $ActionRoot->appendChild($ZoneID);
           $response = $this->fireRDA($xml,$command);
           if(!$response){$templateList[$ID] = array();}
           else{$templateList[$ID] = $response;}
       }
       return $templateList;
    }
    /**
     * Get Bulletin List
     * 
     * Gets the Bulletin information for select zone IDs.  Requires that Zone IDs be set.
     * 
     * @return boolean|array Returns an associative array of Bulletin information organized by Zone ID.
     * Bulletin information contains 'Description', 'AlwaysOn', 'DateTimeOn', 'DateTimeOff', 'CycleTimeOn', 'CycleTimeOff', 'Weekdays' (bitfield), 'WebEnabled', 
     * 'PageType', 'DisplayDuration', 'PreviewImagePath', 'ThumbnailImagePath', 'FullImagePath', 'GUID', 'ObjectType', 'PageStatus', 
     * and a 'Block' array with each value having 'Name' 'Value' and 'Type' attributes.
     * Returns False on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function GetBulletinList(){
       $command =  'GetBulletinList';
       $bulletinList = array();
       if(count($this->ZoneID) === 0){
           $this->LastError = "GetBulletinList() requires that at least one ZoneID be set";
           return 0;
       }
        foreach($this->ZoneID as $ID){
            $xml = new DOMDocument('1.0','utf-8');
           $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
           $xml->appendChild($root);

           $ActionRoot = $xml->createElement($command);
           $root->appendChild($ActionRoot);
           $UserName = $xml->createElement('UserName',$this->UserName);
           $ActionRoot->appendChild($UserName);
           $Password = $xml->createElement('Password',$this->Password);
           $ActionRoot->appendChild($Password);

           $ZoneID = $xml->createElement('ZoneID',$ID);
           $ActionRoot->appendChild($ZoneID);
           $response = $this->fireRDA($xml,$command);
           if(!$response){$bulletinList[$ID] = array();}
           else{$bulletinList[$ID] = $response;}
       }
       return $bulletinList;
    }
    /**
     * Get Picture List
     * 
     * Gets the image media information for select zone IDs.  Requires that Zone IDs be set.
     * 
     * @return boolean|array Returns an associative array of Image information organized by Zone ID
     * Image information contains 'Name', 'Value' (GUID), 'PreviewImagePath', 'TinyImagePath', 'ThumbnailImagePath', and 'FullImagePath'.
     * Returns False on failure.
     * Upon failure an error message will be available in getLastError();
     */
    public function GetPictureList(){
        $command =  'GetPictureList';
       $pictureList = array();
       if(count($this->ZoneID) === 0){
           $this->LastError = "GetPictureList() requires that at least one ZoneID be set";
           return 0;
       }
        foreach($this->ZoneID as $ID){
            $xml = new DOMDocument('1.0','utf-8');
           $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
           $xml->appendChild($root);

           $ActionRoot = $xml->createElement($command);
           $root->appendChild($ActionRoot);
           $UserName = $xml->createElement('UserName',$this->UserName);
           $ActionRoot->appendChild($UserName);
           $Password = $xml->createElement('Password',$this->Password);
           $ActionRoot->appendChild($Password);

           $ZoneID = $xml->createElement('ZoneID',$ID);
           $ActionRoot->appendChild($ZoneID);
           $response = $this->fireRDA($xml,$command);
           if(!$response){$videoList[$ID] = array();}
           else{$pictureList[$ID] = $response;}
       }
       return $pictureList;
    }
    /**
     * Get Picture List
     * 
     * Gets the image media information for select zone IDs.  Requires that Zone IDs be set.
     * 
     * @return boolean|array Returns an associative array of Image information organized by Zone ID
     * Image information contains 'Name', 'Value' (GUID), 'PreviewImagePath', 'TinyImagePath', and 'ThumbnailImagePath'.
     * Returns False on failure. NOTE: this will return an empty array if no video media is present.
     * Upon failure an error message will be available in getLastError();
     */
    public function GetVideoList(){
         $command =  'GetVideoList';
       $videoList = array();
       if(count($this->ZoneID) === 0){
           $this->LastError = "GetVideoList() requires that at least one ZoneID be set";
           return 0;
       }
        foreach($this->ZoneID as $ID){
            $xml = new DOMDocument('1.0','utf-8');
           $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
           $xml->appendChild($root);

           $ActionRoot = $xml->createElement($command);
           $root->appendChild($ActionRoot);
           $UserName = $xml->createElement('UserName',$this->UserName);
           $ActionRoot->appendChild($UserName);
           $Password = $xml->createElement('Password',$this->Password);
           $ActionRoot->appendChild($Password);

           $ZoneID = $xml->createElement('ZoneID',$ID);
           $ActionRoot->appendChild($ZoneID);
           $response = $this->fireRDA($xml,$command);
           if(!$response){$videoList[$ID] = array();}
           else{$videoList[$ID] = $response;}
       }
       return $videoList;
    }
    /**
     * Get Player Status
     * 
     * Gets the current status of every player that has connected to the server.
     * 
     * @return boolean|array Returns an array of player statuses for the server.
     * Player status consists of 'HostName', 'HostAddress', HardwareID, 'VersionStatus', 'PlayerVersion', 'CheckinStatus', 'LastCheckinUTC', and 'SubscribedChannelName'.
     */
    public function GetPlayerStatus(){
       $command =  'GetPlayerStatus';
       
       $xml = new DOMDocument('1.0','utf-8');
       $root = $xml->createElementNS('http://www.trms.com/CarouselRemoteCommand','CarouselCommand');
       $xml->appendChild($root);

       $ActionRoot = $xml->createElement($command);
       $root->appendChild($ActionRoot);
       $UserName = $xml->createElement('UserName',$this->UserName);
       $ActionRoot->appendChild($UserName);
       $Password = $xml->createElement('Password',$this->Password);
       $ActionRoot->appendChild($Password);
       
       return $this->fireRDA($xml,$command); 
    }
    
    private function fireRDA($xml,$command = null){
        $xml->formatOutput = true;
        $XML = $xml->saveXML();
        // var_dump($XML);
        
        try{
            $socket = socket_create(AF_INET, SOCK_STREAM, SOL_TCP);
            if($socket === false){ 
                $this->LastError = "Socket Create Error: ".socket_strerror( socket_last_error() )."\r\n";
                return 0;
            }
            if(! socket_connect($socket, $this->HostIP, $this->Port) ){
                $this->LastError = "Socket Connect Error: ".socket_strerror( socket_last_error($socket) )."\r\n";
                return 0;
            }
            if(! socket_write($socket, $XML, strlen($XML)) ){
                $this->LastError = "Socket Write Error: ".socket_strerror( socket_last_error($socket) )."\r\n";
                return 0;
            }
            $response = '';
            do{
                $buffer = socket_read($socket, 1024);
                $response .= $buffer;
            }
            while(strpos($buffer,'</CarouselResponse>') === FALSE);
            socket_close($socket);
            $response = $this->processResponse($response,$command);
            
            return $response;
        }
        catch(Exception $ex){
            $this->LastError .= "Excpetion: ".$ex->getMessage();
        }
    }
    
    private function processResponse($response,$command){
        //var_dump($response);
        $response = mb_convert_encoding($response,"UTF-16","UTF-8");
        
        $Sxml = simplexml_load_string($response);
        
        if((string)$Sxml->Result === 'Error'){
            $this->LastError = "Carousel Error: ".$Sxml->Description;
            return 0;
        }
        
        else if((string)$Sxml->Result === 'Success'){
            
            if($command === 'DeleteAllUserPages'|| $command === 'ArchivePage' || $command === 'DeletePage' || $command === 'DeactivateAllAlertPages'||$command === 'SetPageStatus'||$command === 'ChangePageStatus'){
                return 1;
            }
            
            if($command === 'CreatePage'||$command === 'UpdatePage'||$command === 'CreateCrawl'||$command === 'UpdateCrawl'){
                $guids = (array)$Sxml->GUID;
                return $guids;
            }
            
            if($command === 'GetBulletinList'){
                $array = json_decode(json_encode((array) $Sxml),1);
                if(isset($array['Bulletin'])){
                    return $array['Bulletin'];
                }
                else return array(); 
            }
            
            if($command === 'GetTemplateList'){
                $array = json_decode(json_encode((array) $Sxml),1);
                if(isset($array['Template'])){
                    return $array['Template'];
                }
                else return array(); 
            }
            
            if($command === 'GetZoneList'){
                $array = json_decode(json_encode((array) $Sxml),1);
                if(isset($array['Zone'])){
                    return $array['Zone'];
                }
                else return array();
            }
            if($command === 'GetPictureList'){
                $array = json_decode(json_encode((array) $Sxml),1);
                if(isset($array['Media'])){
                    return $array['Media'];
                }
                else return array();
            }
            if($command === 'GetVideoList'){
                $array = json_decode(json_encode((array) $Sxml),1);
                if(isset($array['Media'])){
                    return $array['Media'];
                }
                else return array();
            }
            if($command === 'GetPlayerStatus'){
                $array = json_decode(json_encode((array) $Sxml),1);
                if(isset($array['PlayerStatusList'])){
                    return $array['PlayerStatusList']['PlayerStatus'];
                }
                else return array();
            }
            
        }
        else{
            echo "Unhandled: ";
            
        }
        
    }
    
    public function getLastError(){
        return $this->LastError;
    }
}

