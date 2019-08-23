;;;; cl-xlsx.lisp

(in-package #:cl-xlsx)

;; From Carlos Ungil
;; modified by Gwang-Jin Kim
(defun list-entries (file)
  "Internal use, gets entries inside of ZIP/XLSX files."
  (zip:with-zipfile (zip file)
    (let ((entries (zip:zipfile-entries zip)))
      (when entries
	(loop for k being the hash-keys of entries
	      collect k))))) ;; works

;; From Carlos Ungil
(defun get-entry (name zip)
  "Internal use, get content of entry inside the ZIP/XLSX file."
  (let ((entry (zip:get-zipfile-entry name zip)))
    (when entry
      (xmls:parse (babel:octets-to-string
		   (zip:zipfile-entry-contents entry)))))) ;; structs

;; From Gwang-Jin Kim
(defmacro with-open-xlsx ((content-var xml excel-file) &body body)
  "Unzips & parses xml file and binds the parsed result
   to the variable/symbol given for `content-var`.
   In the body part, thus the parsed file content can be referred to
   using the specified symbol/variable."
  (destructuring-bind
      ((content-var xml excel-file))
      `((,content-var ,xml ,excel-file)) ;; (for a nicer looking macro call)
    (let ((zip (gensym)))
      `(zip:with-zipfile (,zip ,excel-file)
	 (let ((,content-var (get-entry ,xml ,zip)))
	   ,@body))))) ;; structs

(defun once-flatten (lst)
  "Return lst just once flattened."
  (cond ((null lst) lst)
	((atom (car lst)) (cons (car lst) (once-flatten (cdr lst))))
	(t (append (car lst) (once-flatten (cdr lst)))))) ; works!

(defun extract-sub-tags (tag-sign tags)
  "Returns the tag-sign matching subtags in a flattened list."
  (once-flatten
   (mapcar #'(lambda (tag)
		(xmls:xmlrep-find-child-tags tag-sign tag))
	   tags)))

(defun collect-extract-exprs (tags acc)
  "Sequencially select tags and flatten inbetween."
  (cond ((null tags) acc)
	(t (collect-extract-exprs (cdr tags)
				  (extract-sub-tags (car tags) acc)))))

;; From Gwang-Jin Kim
(defun select-tags-xlsx (excel-path xml tags) ;; select-xlsx-tags
  "Return tags matching the tags in given hierarchic order.
   Unzips xlsx file (excel-path) and parses xml file before.
   Flattens the results inbetween, thus, output is a plain list of found
   tag objects (list structures defined by the :xmls package)."
  (with-open-xlsx (content xml excel-path)
    (collect-extract-exprs tags (list content))))

;; From Gwang-Jin Kim
(defun select-tags-xmlrep (xmlrep tags)
  "Return xmlrep-tags matching the tags sequentially. 
   Similar to `selet-tags-xlsx`,
   but it does not start with a file,
   but a xmlrep parsed node-object (xmlrep)."
  (collect-extract-exprs tags (list xmlrep)))

(defun attr-val (tag attr)
  "Convenience function."
  (xmls:xmlrep-attrib-value attr tag))

;; From Carlos Ungil
;; modified by Gwang-Jin Kim

(defun get-relationships (xlsx-file)
  "Return relation ships of the excel file. Not for .ods!"
  (let ((relations (select-tags-xlsx xlsx-file
				     "xl/_rels/workbook.xml.rels"
				     '(:relationship))))
    (loop for rel in relations
	  collect (cons (attr-val rel "Id")
			(concatenate 'string "xl/" (attr-val rel "Target"))))))

;; From Carlos Ungil
;; rewritten by Gwang-Jin Kim

(defun get-number-formats (xlsx-file)
  (let* ((formats (select-tags-xlsx xlsx-file "xl/styles.xml"
				   '(:numFmts :numFmt)))
	 (format-codes (loop for fmt in formats
			     collect (cons (parse-integer
					    (attr-val fmt "numFmtId"))
					   (attr-val fmt "formatCode"))))
	 (styles (select-tags-xlsx xlsx-file
				   "xl/styles.xml"
				   '(:cellXfs :xf))))
    (loop for style in styles
	  collect (let ((fmt-id (parse-integer
				 (attr-val style "numFmtId"))))
		    (cons fmt-id
			  (if (< fmt-id 164)
			      :built-in
			      (cdr (assoc fmt-id format-codes))))))))
    
;; From Carlos Ungil
;; modified by Gwang-Jin Kim

(defun column-and-row (colrow)
  (let ((column))
    (loop for char across colrow
	  for pos from 0
	  while (alpha-char-p char) collect char into column
	  finally (cons (intern (coerce column 'string)
				"KEYWORD")
			(parse-integer colrow
				       :start pos)))))

;; From Carlos Ungil

(defun excel-date (int)
  (apply #'format nil "~D-~2,'0-~2,'0D"
	 (reverse
	  (subseq
	   (multiple-value-list
	    (decode-universal-time (* 24
				      60
				      60
				      (- int 2))))
	   3 6))))

;; From Carlos Ungil
;; rewritten by Gwang-Jin Kim

(defun list-sheets (file)
  "Retrieves the id and name of the worksheet in the .xlsx/.xlsm file."
  (let ((sheets (select-tags-xlsx file "xl/workbook.xml" '(:sheets :sheet))))
    (loop for sheet in sheets
	  with rels = (get-relationships file)
	  for sheet-id   = (attr-val sheet "sheetId")
	  for sheet-name = (attr-val sheet "name"   )
	  for sheet-rel  = (attr-val sheet "id"     )
	  collect (list (parse-integer sheet-id)
			sheet-name
			(cdr (assoc sheet-rel rels :test #'string=))))))

;; From Carlos Ungil
;; rewritten by Gwang-Jin Kim

(defun sheet-address (file sheet)
  "Return inner xml address of an excel sheet."
  (let* ((sheets (list-sheets file))
	 (entry-name
	   (cond ((and (null sheet) (= 1 (length sheets)))
		  (third (car sheets)))
		 ((stringp sheet)
		  (third (find sheet sheets :key #'second
					    :test #'string=)))
		 ((numberp sheet)
		  (third (find sheet sheets :key #'first))))))
    (unless entry-name
      (error "specify one of the following sheet ids or names: ~{~&~{~S~^~5T~}~}"
	     (loop for (id name) in sheets
		   collect (list id name))))
    entry-name))

;; From Gwang-Jin Kim
(defun begins-with-p (str substring)
  "String begins with substring?"
  (and (>= (length str) (length substring))
       (string= substring (subseq str 0 (length substring)))))

;; From Gwang-Jin Kim
(defun app-type (file)
  "Return the type of an .xlsx or .ods file."
  (let ((entries (list-entries file)))
    (flet ((extract-app-name (mode)
	     (let* ((file-is-ods (string= mode "ods"))
		    (xml  (if file-is-ods "meta.xml" "docProps/app.xml"))
		    (tags (if file-is-ods '(:meta :generator) '(:Application))))
	       (xmls:xmlrep-string-child ;; crucial for struct!
		(car (select-tags-xlsx file xml tags)))))
	   (is-in-p (string string-list)
	     (member string string-list :test #'string=)))
      (cond ((and (is-in-p "meta.xml" entries)
		  (begins-with-p (extract-app-name "ods") "LibreOffice"))
	     "ods-libreoffice")
	    ((and (is-in-p "docProps/app.xml" entries)
	          (begins-with-p (extract-app-name "xlsx") "LibreOffice"))
	     "xlsx-libreoffice")
	    ((and (is-in-p "docProps/app.xml" entries)
		  (string= (extract-app-name "xlsx") "Microsoft Excel"))
	     "xlsx-microsoft"))))) ;; works!
;; the `car` unpacks the list around the single tag

;; from Carlos Ungil
;; modified by Gwang-Jin Kim
(defun get-unique-strings (xlsx-file)
  "Return unique strings - necessary for parsing excel data."
  (let ((tags (select-tags-xlsx xlsx-file "xl/sharedStrings.xml" '(:si :t))))
    (if (string= (app-type xlsx-file) "xlsx-microsoft")
	(mapcar #'xmls:xmlrep-string-child tags)
        (loop for tag in tags
	      collect (if (equal (xmls:node-attrs tag) '(("space" "preserve")))
			  (xmls:xmlrep-string-child tag)
			  " "))))) ;; corrected by Gwang-Jin Kim 18-09-15

;; (defun get-unique-strings-windows (xlsx-file)
;;   (let ((tags (cl-xlsx:select-tags-xlsx xlsx-file "xl/sharedStrings.xml" '(:si :t))))
;;     (mapcar #'(lambda (x) (third x)) tags))) ;; works!

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; read-in .ods file cell contents as strings
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun process-table-cell (table-cell)
  "Return table-cell content as string."
  (xmls:xmlrep-string-child
   (xmls:xmlrep-find-child-tag :p table-cell)))

(defun process-table-row (table-row)
  "Return list of text values in cells."
  (let ((cell (xmls:xmlrep-find-child-tags "table-cell" table-row)))
    (mapcar #'process-table-cell cell)))

(defun process-table-rows-ods (table-rows)
  "Return list of list of table-row contents as strings."
  (mapcar #'process-table-row table-rows))


(defun read-ods (ods-file)
  "Read all sheets of an ods-file into a list of lists and strings.
   The table contents are list of lists. (row-lists)
   Each sheet is a list. And the entire result is a list of sheets."
  (let ((inner-files (list-entries ods-file)))
    (when (member "content.xml" inner-files :test #'string=)
      (let* ((table-tags (select-tags-xlsx ods-file
					   "content.xml"
					   '(:body :spreadsheet :table)))
	     (tables-as-row-tags (mapcar #'(lambda (sheet-tag)
					     (select-tags-xmlrep
					      sheet-tag
					      '(:table-row)))
					 table-tags)))
	(mapcar #'process-table-rows-ods tables-as-row-tags)))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; read-in .xlsx files as strings
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun extract-val-xlsx-cell (cell-tag)
  (xmls:xmlrep-string-child
   (xmls:xmlrep-find-child-tag :v cell-tag)))

(defun process-table-cell-xlsx (cell-tag unique-strings)
  (let ((val (extract-val-xlsx-cell cell-tag))
	(attribs (mapcar #'car (xmls:xmlrep-attribs cell-tag))))
    (if (and (member "t" attribs :test #'string=)
	     (equalp (attr-val cell-tag "t") "s"))
	(elt unique-strings (parse-integer val))
	val))) ;; else string!

(defun process-table-row-xlsx (row-tag unique-strings)
  (let ((cells (select-tags-xmlrep row-tag '("c"))))
    (mapcar #'(lambda (table-cell)
		(process-table-cell-xlsx table-cell unique-strings))
	    cells)))

(defun process-table-rows-xlsx (table-rows unique-strings)
  (mapcar #'(lambda (table-row)
	      (process-table-row-xlsx table-row unique-strings))
	  table-rows)) ;; works!!


(defun select-sheet-addresses (inner-files)
  (remove-if-not #'(lambda (s) (begins-with-p s "xl/worksheets/"))
		 inner-files))

(defun read-xlsx (xlsx-file)
  (let* ((inner-files (list-entries xlsx-file))
	 (sheet-addresses (select-sheet-addresses inner-files))
	 (unique-strings (get-unique-strings xlsx-file))
	 (sheet-row-lists
	   (loop for sheet-address in sheet-addresses
		 for sheet-rows = (select-tags-xlsx xlsx-file
						    sheet-address
						    '(:sheetData :row))
		 collect (process-table-rows-xlsx sheet-rows unique-strings))))
    (nreverse sheet-row-lists)))


;; the previous versions worked until very recently xmls changed to
;; represent everything as struct.
;; the functions which used list-extraction functions created problems
;; mainly process-cell functions...

;; thus, always, one should abstract such processes
;; and use the abstracted versions (xmls:xmlrep- ...)
;; in the old package xmls:xmlrep- functions worked on the list structures
;; of tags.
;; when extracting strings, I used car, caddar and such functions.

;; still a problem is that xmlrep-find-child-tags -> collects into a list
;; also my functions to select tags collect into a list.
;; and there, I have to unpack when I want to apply
;; xmls:xmlrep-string-child to extract the string ...

;; anyhow, this story showed me, how important it is to use abstractions ...

;; the best is to write an own xml parser (or copy the old one)
;; and 'freeze' it.







;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(ql:quickload :cxml)
(ql:quickload :zip)
;; (ql:quickload :cl-xlsx)


(defun source-entry (xml xlsx)
  "Get content xml file inside xlsx."
  (zip:with-zipfile (zip xlsx)
    (let ((entry (zip:get-zipfile-entry xml zip)))
      (when entry
	(cxml:make-source (babel:octets-to-string (zip:zipfile-entry-contents entry)))))))

;;; extractor functions for sax-parsed cell

(defun sax-cell-value (sax)
  "Return value of sax-parsed cell tag (in child tag :v)."
  (car (last (car (last sax)))))

(defun sax-cell-attrs (sax)
  "Return attribute list of a sax-parsed cell tag."
  (cadr sax))

(defun sax-cell-type (sax)
  "Return type of sax-parsed excel cell tag (attribute 't')."
  (cadr (assoc "t" (sax-cell-attrs sax) :test #'equal)))

(defun sax-cell-pos (sax)
  "Return position information of sax-parsed excel cell tag (attribute 'r')."
  (cadr (assoc "r" (sax-cell-attrs sax) :test #'equal)))

(defun process-sax-cell (sax unique-strings)
  "Return value of sax-parsed excel cell tag. Checks type 't'
   and if string 's', looks up from unique-strings the right string.
   If numeric 'n', then parses it using lisp-reader.
   Otherwise return string." 
  (let ((val-type (sax-cell-type sax)))
    (cond ((equalp val-type "s") (elt unique-strings (parse-integer (sax-cell-value sax))))
	  ((equalp val-type "n") (with-input-from-string (in (sax-cell-value sax))
				   (read in)))
	  (t (sax-cell-value sax)))))


;;;;;;;;;;;;;;;;;;;;;;;;;;
;; parse sheet
;;;;;;;;;;;;;;;;;;;;;;;;;;


;; the final correct version!
(defun get-unique-strings (xlsx)
  "Return all unique strings from xlsx file."
  (klacks:with-open-source (src (source-entry "xl/sharedStrings.xml" xlsx))
    (loop for key = (klacks:peek src)
	  while key
	  nconc (case key
		  (:start-element
		   (if (equal (klacks:current-qname src) "t")
		       (list (let ((x (caddr (klacks:serialize-element src (cxml-xmls:make-xmls-builder)))))
			       (if (null x)
				   "" ;; if none available return empty string!
				   x)))
		       nil))
		  (otherwise nil))
       do (klacks:consume src))))

(defun parse-xlsx-sheet (sheet-xml xlsx)
  "Return parsed content for a given sheet-xml address in a xlsx-fpath."
  (klacks:with-open-source (s (source-entry sheet-xml xlsx))
    (let ((unique-strings (cl-xlsx:get-unique-strings xlsx)))
      (loop for key = (klacks:peek s)
	    while key
	    nconcing (case key
		       (:start-element
			(let ((tag-name (klacks:current-qname s)))
			  (cond ((equal tag-name "row")
				 (loop for key = (klacks:peek s)
				    for consumed = nil
				    while key
				    nconcing (case key
					       (:start-element
						(cond ((equal (klacks:current-qname s) "c")
						       (setf consumed t)
						       (list (process-sax-cell
							      (klacks:serialize-element
							       s
							       (cxml-xmls:make-xmls-builder))
							      unique-strings)))
						      (t
						       (setf consumed nil)
						       nil)))
					       (:end-element
						(if (equal (klacks:current-qname s) "row")
						    (return (list res)))))
				    into res
				    do (unless consumed
					 (klacks:consume s))))
				(t nil))))
		       (otherwise nil))
	 do (klacks:consume s))))) ;; works!


;;;


(defun starts-with-p (str substring)
  "String starts/begins with substring?"
  (let ((str-len (length str))
	(sub-len (length substring)))
    (and (>= str-len sub-len)
	 (string= substring (subseq str 0 sub-len)))))

(defun ends-with-p (str substring)
  "String ends with substring?"
  (let ((str-len (length str))
	(sub-len (length substring)))
    (and (>= str-len sub-len)
	 (string= substring (subseq str (- str-len sub-len) str-len)))))
				    
(defun inner-files (xlsx)
  "List all innter addresses in xlsx."
  (zip:with-zipfile (inzip xlsx)
    (let ((entries (zip:zipfile-entries inzip)))
      (when entries
	(loop for k being the hash-keys of entries
	     collect k)))))

#|
(defun sheets (xlsx)
  "Return sheet informations as list of lists (name sheet-number sheetaddress."
  (let ((sheet-addresses (remove-if-not #'(lambda (x) (and (starts-with-p x "xl/worksheets/")
							   (ends-with-p x ".xml")))
					(inner-files xlsx))))
    (klacks:with-open-source (src (source-entry "xl/workbook.xml" xlsx))
      (loop for key = (klacks:peek src)
	 while key
	 nconc (case key
		 (:start-element
		  (if (equal (klacks:current-qname src) "sheet")
		      (list (let* ((sax (klacks:serialize-element src (cxml-xmls:make-xmls-builder)))
				   (attributes (cdadr sax)))
			      (list (cadr (assoc "name" attributes :test #'equal))
				    (parse-integer (cadr (assoc "sheetId" attributes :test #'equal)))
				    (pop sheet-addresses))))
		      nil))
		 (otherwise nil))
	 do (klacks:consume src))))) ;; works
|#

(defun sheets (xlsx)
  "Return sheet informations as list of lists (name sheet-number sheetaddress)."
  (klacks:with-open-source (src (source-entry "xl/workbook.xml" xlsx))
    (loop for key = (klacks:peek src)
       while key
       nconc (case key
	       (:start-element
		(if (equal (klacks:current-qname src) "sheet")
		    (list (let* ((sax (klacks:serialize-element src (cxml-xmls:make-xmls-builder)))
				 (attributes (cdadr sax)))
			    (list (cadr (assoc "name" attributes :test #'equal))
				  (parse-integer (cadr (assoc "sheetId" attributes :test #'equal)))
				  (concatenate 'string
					       "xl/worksheets/sheet"
					       (cadr (assoc "sheetId" attributes :test #'equal))
					       ".xml"))))
		    nil))
	       (otherwise nil))
       do (klacks:consume src))))

(defun sheet-names (xlsx)
  "List sheet names in xlsx file."
  (mapcar #'car (sheets xlsx)))

(defun sheet-address (sheet xlsx)
  "Return sheet address inside xlsx file when sheet name or index given as input."
  (typecase sheet
    (string (caddr (assoc sheet (sheets xlsx) :test #'string=)))
    (integer (cadr (assoc sheet (mapcar #'cdr (sheets xlsx)))))))

(defun parse-xlsx (xlsx)
  "Parse every sheet of xlsx and return as alist (sheet-name sheet-content-as-list)."
  (let ((sheet-names (sheet-names xlsx)))
    (mapcar #'(lambda (sheet)
		(list sheet (parse-xlsx-sheet sheet xlsx)))
	    sheet-names)))

(defun app-name (xlsx)
  "Return app-name of xlsx or ods/ots file."
  (let ((inner-files (inner-files xlsx)))
    (cond ((member "meta.xml" inner-files :test #'string=)
	   (let ((src (source-entry "meta.xml" xlsx)))
	     (klacks:find-element src "generator")
	     (car (last (klacks:serialize-element src (cxml-xmls:make-xmls-builder))))))
	  ((member "docProps/app.xml" inner-files :test #'string=)
	   (let ((src (source-entry "docProps/app.xml" xlsx)))
	     (klacks:find-element src "Application")
	     (car (last (klacks:serialize-element src (cxml-xmls:make-xmls-builder))))))
	  (t
	   "")))) ; works!

(defun app-type (xlsx)
  "Return type of xlsx file - ods or odt included."
  (let ((inner-files (inner-files xlsx)))
    (cond ((starts-with-p (app-name xlsx) "LibreOffice")
	   (if (member "meta.xml" inner-files :test #'string=)
	       "ods-libreoffice"
	       "xlsx-libreoffice"))
	  ((starts-with-p (app-name xlsx) "Microsoft Excel")
	   "xlsx-microsoft"))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; read ods
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun parse-ods-cell (sax)
  "Return string or number according to type of ods cell sax."
  (let ((type (car (last (first (second sax)))))
        (value (car (last (third sax)))))
    (cond ((equalp type "float") (with-input-from-string (in value)
				   (read in)))
	  ((equalp type "string") value)
	  (t value))))

(defun list-structure (xml xlsx)
  (klacks:with-open-source (s (source-entry xml xlsx))
    (loop for key = (klacks:peek s)
	  while key
	  do (case key
	       (:start-element (format t "~A {" (klacks:current-qname s)))
	       (:end-element (format t "}")))
	     (klacks:consume s))))

(defun number-columns-repeated-cell-p (table-cell-sax)
  (equal (car (caaadr table-cell-sax)) "number-columns-repeated"))

(defun flatten (l &key (acc '()))
  (cond ((null l) acc)
	((atom (car l)) (flatten (cdr l) :acc (append acc (list (car l)))))
	(t (flatten (cdr l) :acc (append acc (flatten (car l)))))))

(defun all-nil-row-p (l)
  (every #'null (flatten l)))

(defun parse-ods (ods)
  "Return parsed content for a given sheet-xml address in a xlsx-fpath."
  (klacks:with-open-source (s (source-entry "content.xml" ods))
    (loop for key = (klacks:peek s)
	  while key
	  nconcing (case key
		     (:start-element
		      (cond ((equal (klacks:current-qname s) "table:table")
			     (loop for key1 = (klacks:peek s)
				   while key1
				   nconcing (case key1
					      (:start-element
					       (cond ((equal (klacks:current-qname s) "table:table-row")
						      (loop for key2 = (klacks:peek s)
							    for consumed = nil
							    while key2
							    nconcing (case key2
								       (:start-element
									(cond ((equal (klacks:current-qname s) "table:table-cell")
									       (setf consumed t)
									       (let ((sax (klacks:serialize-element s (cxml-xmls:make-xmls-builder))))
										 (if (number-columns-repeated-cell-p sax)
										     nil
									             (list (parse-ods-cell sax))))) ; this corrected the final NIL's in row
									      (t
									       (setf consumed nil)
									       nil)))
								       (:end-element
									(if (equal (klacks:current-qname s) "table:table-row")
									    (return (if (all-nil-row-p inner-res) ; this is a hack to get rid of tags at end of sheet
											nil ; remove NIL-only row constructs
											(list inner-res)))))
								       (otherwise nil))
							      into inner-res
							    do (unless consumed
								 (klacks:consume s))))
						     (t nil)))
					      (:end-element
					       (if (equal (klacks:current-qname s) "table:table")
						   (return (list res))))
					      (otherwise nil))
				     into res
				   do (klacks:consume s)))
			    (t nil)))
		     (otherwise nil))
	  do (klacks:consume s))))

;;; I want one `read-xlsx` function to handle both, ods and xlsx.
;;; So it needs some introspection.
;;; I should give an associative list back
;;; where from each pair the first element is the sheet name
;;; and the second element the parsed sheet.

;;; Later:
;;; It should have an optional :sheet argument by which one can restrict reading to a sheet or sheets.

;;; There should be a sheet-names function which does the reading of sheet names.

;;; I need a file-type function which returns "ods" or "xlsx". For introspection.

;;;




#|

First part is parsing xlsx and ods files into lists at once.

`list-entries` lists all filenames within xlsx or ods (which are zipped bundle of XML files).

`get-entry` extracts sheets from zip
-> it should be from xlsx/ods directly!

`with-open-xlsx is wrapper um `zip:with-zipfile` where content-var is the get-entried content.
however this abstraction is not very happy one.
Rather one should have a sheet content extractor from xlsx direct.

`once-flatten` is a typical utility function which flattens list just by one level.

`extract-sub-tags returns tag-sign matching subtgs in a flattened list.
`collect-extract-exprs` sequencially extract tags and flattens inbetween.
it is a little like xpath.

`select-tags-xlsx` extracts tag content directly from sheet-address and excelpath.
There should be not sheet-address but sheet-name or sheet-number, because user-friendlier.

`select-tags-xmlrep` not from xlsx directly, but a parsed xml file inside xlsx.
xmlrep = SAX

`attr-val` returns value of attribute in a tag.

`get-reglationships` peeks in xlsx into "xl/_rels/workbook.xml.rels"
and returns sheet id and address.

the basic idea was that through xpath like order in the tags, the collection is more
understandable and easier - but this is not the case.
the klacks-syntax is easy - because to think of everything as stream and just to depict
the last level sometimes helps ...

`get-number-formats` tries to get correct number format in xlsx it seems.

`column-and-row`

`excel-date` tries to get excel dat correcty.

`list-sheets` lists the sheet names in the xlsx file.

`sheet-address` makes out of sheet-number or sheet-name (given for sheet)

`begins-with-p`

`app-type` determines the app type (xlsx or ods)

`get-unique-strings` extracts all unique strings from xlsx-file.



`process-table-cell` processes an ods table cell - just extracts value
`process-table-row` extract 'table-cell's in a 'table-row' tag and applies `process-table-cell` on them.
`process-table-row-ods` applies `process-table-row` on all rows of a table.

`read-ods` lists first the innter-files using list-entries,
tests whehter "content.xml" is amongst them (ods file!)
extract tables, then table-rows for each table
and then process-table-rows-ods on tables as row-tags.
list of sheets


`extract-val-xlsx-cell` 
  extracts value from a xlsx cell
`process-table-cell-xlsx`
  takes `unique-strings` and a celltag and if it is a string, searches from unique-strings
  the correct string, otherwise parses integer, otherwise the string as it is.
`process-table-row-xlsx` 
  retuires `unique-strings` obviously and applies `process-table-cell-xlsx` to each row in table.
`select-sheet-addresses` of inner-files - selects only the sheet-addresses
  (those beginning with "xl/worksheets/")

`read-xlsx`
  extracts all addresses
  gets sheets addreses
  extracts all unique strings
  selects the sheet-row-lists
    by extracting rows for each address and `process-table-rows-xlsx` on them
      using unique-strings
  and nreverses sheet-row-lists before returning

;; maybe all these functions should be made dependent on :cxml and not :xlsx

;;;;;;;;;;;;;;;;;;;;;;;;
;; cxml
;;;;;;;;;;;;;;;;;;;;;;;;

`source-entry` gets content directly from xlsx file
  (using `zip:get-zipfile-entry)
  I can put core - getting from xlsx an xml - to: `get-content` (xml xlsx)

`sax-cell-value` extracts value from sax-cell
`sax-cell-attrs` extracts attributes from sax-cell
`sax-cell-type` extracts type information from sax-cell
`sax-cell-pos` extrcts position information from sax-cell (A1 B1 etc.)
if using sax cell, this can be used from cxml as well as klacks
so the above functions for :xlsx should be replaced by these

`process-sax-cell` looksup whether string or not
  if from strings then looksup right string from unique-strings
  otherwise if number then uses :cl reader
  otherwise just the value
(this is like that above and can replace previous function).

(just the parser should be from :cxml - and be the sax parser)


;; `get-unique-strings` extracts unique strings but in cases that there is nothing
;; it returns empty string

`parse-xlsx-sheet` extracts and parses one xml in xlsx

`starts-with-p` tests just whether a string begins with a string
`ends-with-p` tests the ending

`inner-files` lists the addresses inside an xlsx or ods

`sheets` returns all sheet informations (sheet-name sheet-number sheet-address)
`sheet-names` returns only the names
`sheet-address` returns the address of sheet in xlsx, as long as sheet is either name or number

`parse-xlsx` reads in every sheet
;; there should be one inbetween step - parse-xlsx for only 1 sheet - but sheet given as name or number
;; the address is sth intern

`app-name` determines app name
`app-type` returns ods-libreoffice or xlsx-libreoffice or xlsx-microsoft


`parse-ods-cell` parses value of one ods-cell
`list-structure` is just from the example - to see roughly the structure

`number-columns-repeated-cell-p` is necessary to detect ods-cells without values
`flatten` is similar to alexandria:flatten
`all-nil-row-p` is a hack for removing only NIL element containing rows

`parse-ods` reads in all ods sheets







|#
