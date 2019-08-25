cl-xlsx
=======

An easy-to-use XLSX reader in Common Lisp, which can read .xlsx and .ods files
created by Microsoft Office or LibreOffice, respectively.

The objective of this package was to not to have to convert .xlsx/.ods files
into .csv before reading of data tables in .xlsx. Therefore the .xlsx or .ods
files should be very simple—only one simple table (csv table) per sheet, and no
compex sheet design, and no complex (fused) cells.


Dependencies
------------

[cxml](https://github.com/sharplispers/cxml) is used for SAX-parsing using
streams, while [zip](https://github.com/mcna/zip) and
[babel](https://github.com/cl-babel/babel) are used for handling binary streams.

Be aware, though, that reading very large files is slow—more than 15 seconds for
a 2–5 MB file. The upside, however, are efficient handling of the files due to
streams.


Usage
-----

Clone this repository into the `~/common-lisp/`:

```
cd ~/common-lisp
git clone https://github.com/a1b10/cl-xlsx
```

Quickload it, then define paths:

```lisp
(ql:quickload :cl-xlsx)
(defparameter *ods* #P"/path/to/your/file.ods")
(defparameter *xlsx* #P"/path/to/your/file.xlsx")
```

To read the tables:

```lisp
(defparameter *ods-sheet-contents* (cl-xlsx:read-xlsx *ods*))
(defparameter *xlsx-sheet-contents* (cl-xlsx:read-xlsx *xlsx*))
```


### Return format

Tables are returned in the form of `alist`s:

```
((sheet1-name . sheet1-content)
 (sheet2-name . sheet2-content))
```

The sheet content is a simple list of lists where the cdr is the content
of a row in the table:

```lisp
((row1-element-1 row1-element-2 ... row1-element-k)
 (row2-element-1 row2-element-2 ... row2-element-k)
 ...
 (rowN-element-1 rowN-element-2 ... rowN-element-k))
```

Only numbers are currently correctly parsed into numbers. All other cell values
are returned as strings.

We suggest you to write your own functions to distinguish different formats
(e.g. dates) and to write your own transformation functions for them
(e.g. function to recognize date-strings and converter functions to your
date-representation of choice).


### Retrieve only sheet names

```lisp
(cl-xlsx:sheet-names *ods*)
(cl-xlsx:sheet-names *xlsx*)
```


### Select content of one sheet

By name:

```lisp
(defun select-sheet (sheet-name xlsx-content)
  (cdr (assoc sheet-name xlsx-content :test #'string=)))

(defparameter *sheet-1-content* (select-sheet "Sheet1" *xlsx-content*))

```

By number (beginning from 1):

```lisp
(defparameter *sheet-1-content* (select-sheet 1 *xlsx-content*))
(defparameter *sheet-2-content* (select-sheet 2 *xlsx-content*))
```
