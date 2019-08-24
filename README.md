cl-xlsx
=======

Easy to use xlsx-reader in Common Lisp, which can read .xlsx and .ods files created by MS-Office or Libreoffice btw. Libreoffice, respecitvely.

The objective of this package was to not to have to convert .xlsx/.ods files into .csv before 
reading of data tables in .xlsx.
Therefore the .xlsx or .ods files should be very simple (only one simple table (csv table) per sheet - (no compex sheet design and no complex (fused) cells! Just no complex formatting.).


# Dependencies

`:cxml` and there `klacks` for SAX-parsing using streams. 
Beware reading is very slow for big files (>15 seconds for a 2~5 Mb file!) but this is the cost for memory efficient treatment of the files (as streams).
`:zip` and `:bable` for handling binary streams.

# Usage

Install and define paths:
```
(ql:quickload :cl-xlsx)
(defparameter *ods* #P"/path/to/your/file.ods")
(defparameter *xlsx* #P"/path/to/your/file.xlsx")
```

Read the tables (:cl-xlsx can handle only simple sheets with simple tables in them! The package is yet in a very beginning stage.)
Tables are returned in the form of plists: 
```
'((sheet1-name . sheet1-content) 
     (sheet2-name . sheet2-content))
```
The sheet content is a simple list of list where every inner list is the content of a row in the table: 
```'((row1-element-1 row1-element-2 ... row1-element-k) 
     (row2-element-1 row2-element-2 ... row2-element-k) 
     ... 
     (rowN-element-1 rowN-element-2 ... rowN-element-k))
```
```
(defparameter *ods-sheet-contents* (cl-xlsx:read-xlsx *ods*))
(defparameter *xlsx-sheet-contents* (cl-xlsx:read-xlsx *xlsx*))
```

For inspection of the names of the sheets there is the `sheet-names` function.

```
;; inspect sheet names
(cl-xlsx:sheet-names *ods*)
(cl-xlsx:sheet-names *xlsx*)
```


Select content of one sheet:
```
(defun select-sheet (sheet-name xlsx-content)
  (cdr (assoc sheet-name xlsx-content :test #'string=)))

(defparameter *sheet-1-content* (select-sheet "Sheet1" *xlsx-content*))

```
It is actually not necessary to now the sheet names for the extraction.
You can use simply the sheet number beginning from 1:
```
(defparameter *sheet-1-content* (select-sheet 1 *xlsx-content*))
(defparameter *sheet-2-content* (select-sheet 2 *xlsx-content*))
```

To not to have everytime to re-read (re-parse) the xlsx file, one should save the contents first
into a variable. And use this variable for extractions.

Happy reading of Excel-files!
And happy downstream processing of the read content! ;)

