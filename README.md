Excel Merge
===========

> 🧩 This package is built on top of https://github.com/infostreams/excel-merge, which served as a great starting point.

Merges two or more Excel files into one file, while keeping formatting intact. This package merges the sheets of the individual excel sheets into one without loading the whole excel file into memory. This software works with Excel 2007 (.xlsx and .xlsm) files and can only generate Excel 2007 files as output. The older .xls format is unfortunately not supported.

However, we found that the original package was no longer actively maintained and didn’t fully meet our project’s requirements out of the box. In particular, it lacked support for handling Excel files containing attachments, an essential feature for our workflow. Upon investigation, we found that the issue stemmed from the package not properly copying the `_rels` directory when unzipping and reassembling Excel files, resulting in broken or missing file links. Additionally, our project has a strong emphasis on code reliability and required a well-tested foundation—something the original package did not provide. As a result, we built this package with improved extensibility, proper handling of attachments, and comprehensive unit tests.

Installation
------------

**With composer**

    composer require nzalheart/excel-merge

Use
---

The most basic use of this software looks something like this 

    <?php
      require("vendor/autoload.php");
    
      $files = array("generated_file.xlsx", "tmp/another_file.xlsx");
      
      $merged = new ExcelMerge\ExcelMerge($files);            
      $merged->download("my-filename.xlsm");
      
      // or
      
      $filename = $merged->save("my-directory/my-filename.xlsm");
    ?>


How it works
------------
Instead of trying to keep a mental model of the whole Excel file in memory, this library simply 
operates directly on the XML files that are inside Excel2007 files. The library doesn't 
really understand these XML files, it just knows which files it needs to copy where and how to
modify the XML in order to add one sheet of one Excel file to the other. 

This means that the most memory it will ever use is directly related to how large your largest
worksheet is.

Requirements
------------
This library uses DOMDocument and DOMXPath extensively. Please make sure these extensions are installed.