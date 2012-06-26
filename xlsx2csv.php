<?php 
/**
 * xlsx2csv.php converts .xlsx files to .csv format
 * Released under the GNU/LGPL licences -- David Collins -- June, 2012 
 *  
 * You may freely use, modify or redistribute this script provided this header remains intact
 * Functions derived from online sources are noted inline
 * The included pclzip PHP zip library  is licensed as noted in related files
 *    
 * @title      xlsx2csv.php 
 * @author     David Collins <collidavid@gmail.com>
 * @license    http://www.gnu.org/copyleft/lesser.html  LGPL License 2.1
 * @version    0.2
 * @link       https://github.com/davidcollins/xlsx2csv
 */
 
 
/**
 * Declare the file to be converted as '$file='. 
 * In demo $file declared at index.php
 */

if(!isset($file)){
$file="";};
/**
 * Set $throttle to limit number of rows converted;
 * Leave blank to process entire file.
 * In demo $throttle declared in index.php
 */
if(!isset($throttle)){
$throttle="";};                                                                 
/**
 * Set $cleanup to 1 for debugging or to leave unpacked files on server;
 * Set to 0 or "" to delete unpacked files in production environment
 */
$cleanup ="0";
/**
 * Set $unpack to 1 if files are already unpacked files on server;
 * Set to 0 or "" to unpack files in production environment
 */
$unpack = "0";

/**
 * Assign CSV file name similar to xlsx file;
 */ 
 $newcsvfile  = str_replace(".xlsx",".csv",$file);
 $newcsvfile = str_replace(" ","-",$newcsvfile);
 $newcsvfile = "csv/$newcsvfile";
 if(!is_dir('bin')) {mkdir("bin", 0770);}; 
 if(!is_dir('csv')) {mkdir("csv", 0777);};
/**
 * Use the PCLZip library to unpack the xlsx file to '/bin'
 * PCLZip will create '/bin' or any other directory named in extract()
 * unpack-directory 
 */
if($unpack!="1"){
require_once 'PCLZip/pclzip.lib.php'; 
 $archive = new PclZip($file);
 $list = $archive->extract(PCLZIP_OPT_PATH, "bin"); 
 }
 

function xmlObjToArr($obj) {

/**
 * convert xml objects to array
 * function from http://php.net/manual/pt_BR/book.simplexml.php
 * as posted by xaviered at gmail dot com 17-May-2012 07:00
 * NOTE: return array() ('name'=>$name) commented out; not needed to parse xlsx
 */
        $namespace = $obj->getDocNamespaces(true);
        $namespace[NULL] = NULL;
       
        $children = array();
        $attributes = array();
        $name = strtolower((string)$obj->getName());
       
        $text = trim((string)$obj);
        if( strlen($text) <= 0 ) {
            $text = NULL;
        }
       
        // get info for all namespaces
        if(is_object($obj)) {
            foreach( $namespace as $ns=>$nsUrl ) {
                // atributes
                $objAttributes = $obj->attributes($ns, true);
                foreach( $objAttributes as $attributeName => $attributeValue ) {
                    $attribName = strtolower(trim((string)$attributeName));
                    $attribVal = trim((string)$attributeValue);
                    if (!empty($ns)) {
                        $attribName = $ns . ':' . $attribName;
                    }
                    $attributes[$attribName] = $attribVal;
                }
               
                // children
                $objChildren = $obj->children($ns, true);
                foreach( $objChildren as $childName=>$child ) {
                    $childName = strtolower((string)$childName);
                    if( !empty($ns) ) {
                        $childName = $ns.':'.$childName;
                    }
                    $children[$childName][] = xmlObjToArr($child);
                }
            }
        }
         
        return array(
           // name not needed for xlsx
           // 'name'=>$name,
            'text'=>$text,
            'attributes'=>$attributes,
            'children'=>$children
        );
    } 
    
function my_fputcsv($handle, $fields, $delimiter = ',', $enclosure = '"', $escape = '\\') {
/**
 * write array to csv file
 * enhanced fputcsv found at http://php.net/manual/en/function.fputcsv.php
 * posted by Hiroto Kagotani 28-Apr-2012 03:13
 * used in lieu of native PHP fputcsv() resolves PHP backslash doublequote bug
 * !!!!!! To resolve issues with escaped characters breaking converted CSV, try this:
 * Kagotani: "It is compatible to fputcsv() except for the additional 5th argument $escape, 
 * which has the same meaning as that of fgetcsv().  
 * If you set it to '"' (double quote), every double quote is escaped by itself."
 */
  $first = 1;
  foreach ($fields as $field) {
    if ($first == 0) fwrite($handle, ",");

    $f = str_replace($enclosure, $enclosure.$enclosure, $field);
    if ($enclosure != $escape) {
      $f = str_replace($escape.$enclosure, $escape, $f);
    }
    if (strpbrk($f, " \t\n\r".$delimiter.$enclosure.$escape) || strchr($f, "\000")) {
      fwrite($handle, $enclosure.$f.$enclosure);
    } else {
      fwrite($handle, $f);
    }

    $first = 0;
  }
  fwrite($handle, "\n");
}

$strings = array();  
$dir = getcwd();
$filename = $dir."\bin\xl\sharedstrings.xml";   

/**
 * XMLReader node-by-node processing improves speed and memory in handling large XLSX files
 * Hybrid XMLReader/SimpleXml approach 
 * per http://stackoverflow.com/questions/1835177/how-to-use-xmlreader-in-php
 * Contributed by http://stackoverflow.com/users/74311/josh-davis
 * SimpleXML provides easier access to XML DOM as read node-by-node with XMLReader
 * XMLReader vs SimpleXml processing of nodes not benchmarked in this context, but...
 * published benchmarking at http://posterous.richardcunningham.co.uk/using-a-hybrid-of-xmlreader-and-simplexml
 * suggests SimpleXML is more than 2X faster in record sets ~<500K
 */

$z = new XMLReader;
$z->open($filename);

$doc = new DOMDocument;

$csvfile = fopen($newcsvfile,"w");

while ($z->read() && $z->name !== 'si');
ob_start();

while ($z->name === 'si')
  { 
    // either one should work
    $node = new SimpleXMLElement($z->readOuterXML());
   // $node = simplexml_import_dom($doc->importNode($z->expand(), true));
        
$result = xmlObjToArr($node);   
$count = count($result['text']) ;
   
if(isset($result['children']['t'][0]['text'])){
  
   $val = $result['children']['t'][0]['text'];
  $strings[]=$val;
 
    };                   
    $z->next('si');
    $result=NULL;      
    };
ob_end_flush();
$z->close($filename);

$dir = getcwd();
$filename = $dir."\bin\xl\worksheets\sheet1.xml";    
$z = new XMLReader;
$z->open($filename);

$doc = new DOMDocument;

$rowCount="0";
$doc = new DOMDocument; 
$sheet = array();  
$nums = array("0","1","2","3","4","5","6","7","8","9");
while ($z->read() && $z->name !== 'row');
ob_start();

while ($z->name === 'row')
  {  
    $thisrow=array();

$node = new SimpleXMLElement($z->readOuterXML());
$result = xmlObjToArr($node); 

$cells = $result['children']['c'];
$rowNo = $result['attributes']['r']; 
$colAlpha = "A";

foreach($cells as $cell){

if(array_key_exists('v',$cell['children'])){

$cellno = str_replace($nums,"",$cell['attributes']['r']);

for($col = $colAlpha; $col != $cellno; $col++) {
 $thisrow[]=" ";
 $colAlpha++; 
   };

  if(array_key_exists('t',$cell['attributes'])&&$cell['attributes']['t']='s'){
    $val = $cell['children']['v'][0]['text'];
    $string = $strings[$val] ;
    $thisrow[]=$string;
      } 
    else {
    $thisrow[]=$cell['children']['v'][0]['text'];
      }
    }
    else {$thisrow[]="";};
    $colAlpha++;
  };

$rowLength=count($thisrow);
$rowCount++;
$emptyRow=array();

while($rowCount<$rowNo){
  for($c=0;$c<$rowLength;$c++) {
    $emptyRow[]=""; 
  }

if(!empty($emptyRow)){
    my_fputcsv($csvfile,$emptyRow);
    };
    $rowCount++;
  };

my_fputcsv($csvfile,$thisrow);      

if($rowCount<$throttle||$throttle==""||$throttle=="0")
  {$z->next('row');
    } 
    else 
      {break;};

$result=NULL; };

$z->close($filename);
ob_end_flush(); 

/**
 * Delete unpacked files from server
 */ 
function cleanUp($dir) {
    $tempdir = opendir($dir);
    while(false !== ($file = readdir($tempdir))) {
        if($file != "." && $file != "..") {
             if(is_dir($dir.$file)) {
                chdir('.');
                cleanUp($dir.$file.'/');
                rmdir($dir.$file);
            }
            else
                unlink($dir.$file);
        }
    }
    closedir($tempdir);
}
if($cleanup!="1"){
  cleanUp("bin/");  
};
?>