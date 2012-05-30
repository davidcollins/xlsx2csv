/**
 * xlsx2csv.php converts .xlsx files to .csv format
 * Released under the GNU/LGPL licences -- David Collins -- June, 2012 
 *  
 * You may freely use, modify or redistribute xlsx2csv.php provided this header remains intact
 * Functions derived from online sources are noted inline
 * The included pclzip library is licensed as noted in related files
 *    
 * @title      xlsx2csv.php 
 * @author     David Collins <collidavid@gmail.com>
 * @license    http://www.gnu.org/copyleft/lesser.html  LGPL License 2.1
 * @version    0.1
 * @link       http://
 */
 
 
0 - SUMMARY
=======

    1 - PURPOSE
    2 - REQUIREMENTS
    3 - HOW IT WORKS
    4 - CONFIGURATION
    5 - LIMITATIONS 
    6 - ALTERNATIVES
    7 - FEEDBACK 

1 - PURPOSE
===========

xlsx2csv.php converts files from the Microsoft 2007 xlsx spreadsheet format to csv format. The csv format provides wide interoperability among data-processing systems. The xlsx2csv script provides a low-resource means for server-side translation. The script is intended to avoid memory or time-out limitations associated with other PHP-based translation approaches when handling large xlsx files.

2 - REQUIREMENTS
================

The script has been tested in a PHP5 environment on an Apache server. The included pclzip package must be installed along with xlsx2csv. 

3 - HOW IT WORKS
================

Files in the Microsoft xlsx format comprise a zipped package of xml files. Xlsx2csv uses PhpConcept's open-source pclzip libarary to unzip xlsx files into a temporary /bin directory. 

Xlsx2csv then uses PHP's XMLReader libraries to read XML components of the xlsx format line by line. The iterative approach avoids resource consumption associated with loading large files into memory. XML nodes are then processed by PHP's SimpleXML library. 

Xlsx2csv reads only two files from the unzipped xlsx package: /xl/sheet1.xml and sharedStrings.xml. SharedStrings contains strings referenced in sheet1.xml.

Translated files are stored in a /csv directory, then files unpacked to the /bin directory are deleted. As packaged for release, an index.php file provides for upload of xlsx files via a Web browswer and returns converted files to the user's browser for download. 

The index.php page is provided solely as a demonstration to give developers a means to test the package and to discover functionality for installation in various contexts. All core functionality resides in xlsx2csv.php file and in the pclzip.lib.php file.
  
Sampleformlist.xlsx -- a small xlsx file -- is included for testing and demonstration purposes. 

4 - CONFIGURATION
=================

Installed on a server, the demo package can open and translate xlsx files with no other configuration. In a production environment, the only required configuration is to declare a file to be translated.

$file at the top of xlsx2csv.php declares the file to be translated. 

Three other configurable variables provide development options: 

$throttle can be set to limit the number of lines processed in large files. 

$cleanup can be toggled off to allow unpacked files to remain in the /bin directory. 

$unpack can be toggled off to allow xlsx2csv to read existing xlsx files in the /bin directory that would otherwise be overwritten by newly unzipped files.

A modified my_fputcsv function provides access to familiar PHP fputcsv parameters (delimiter and string-enclosure options) and an additional parameter that can be modified to work around a known PHP bug related to escaping backslash-doublequote sequences.

5 - LIMITATIONS
===============

As configured for release, xlsx2csv reads only the first sheet of multiple-page workbooks. The resulting csv format preserves no formatting information. As typical of csv files, numerals including dates and currency are typed only as plain-text strings. What you see is what you get.

The original version of this script released in June, 2012 has not been extensively tested against a wide variety of xlsx documents. The original version was tested against documents of more than 10,000 lines. The script was written with scalability in mind, but has not been tested to determine the maximum size of files that can be processed. 

For handling very large files, PHP memory and time-out limits can sometime be increased via php.ini files. 

The script is provided as is, and users are solely responsible for providing whatever security controls are necessary to validate and filter files processed by this script. The script includes a clean-up function that deletes temporary files from a server after a csv file is written. No user should deploy or modify any script -- including this one -- that deletes files from a server unless they are familiar with problems that can arise, and have otherwise provided means for data-preservation.

xlsx2csv is NOT a desktop spreadsheet-processing software.

6 - ALTERNATIVES
================

Numerous programs and scripting alternatives are available for conversion of the xlsx format. The open-source PHPExcel libraries offer extensive tools for server-side management of various Microsoft spreadsheet formats. OpenOffice provides a free open-source desktop environment to read and translate xlsx documents. And, of course, Microsoft's software products provide a wide range of functionality   .

7 - FEEDBACK
============

The author welcomes comments, feedback, suggestions, bug-patches and modifications. The xlsx package is released as-is with no guarantee it is suitable for any purpose. The  package was originally developed to provide server-side xlsx-to-csv translation for automated analysis of public information in support of a journalistic mission. The author invites users to submit reports of any other purposes for which the package may be found useful. 

Contact David Collins at collidavid@gmail.com

 