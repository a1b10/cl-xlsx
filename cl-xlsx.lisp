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
	      collect k)))))

;; From Carlos Ungil
(defun get-entry (name zip)
  "Internal use, get content of entry inside the ZIP/XLSX file."
  (let ((entry (zip:get-zipfile-entry name zip)))
    (when entry
      (xmls:parse (babel:octets-to-string
		   (zip:zipfile-entry-contents entry))))))

;; From Gwang-Jin Kim
(defmacro with-open-xlsx ((content-var xml excel-file) &body body)
  "Unzips & parses xml file and binds to variable given for `content-var`.
   In the body part, access file content using this variable."
  (destructuring-bind ((content-var xml excel-file))
      `((,content-var ,xml ,excel-file)) ;; (for a nicer looking macro call)
    (let ((zip (gensym)))
      `(zip:with-zipfile (,zip ,excel-file)
	 (let ((,content-var (get-entry ,xml ,zip)))
	   ,@body)))))

;; (defun select-child-tags (excel-file xml tag)
;;   "Returns all chidren tags matching given tag in the xml file of the xlsx file."
;;   (zip:with-zipfile (zip excel-file)
;;     (let ((content (get-entry xml zip)))
;;       (xmls:xmlrep-find-child-tags tag content)))) ; works

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
   Unzip xlsx file (excel-path) and parses xml file before.
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
  "Return relation ships of the excel file."
  (let ((relations (select-tags-xlsx xlsx-file
				     "xl/_rels/workbook.xml.rels"
				     '(:relationship))))
    (loop for rel in relations
	  collect (cons (attr-val rel "Id")
			(attr-val rel "target")))))

;; From Carlos Ungil
;; rewritten by Gwang-Jin Kim

(defun get-number-formats (xlsx-file)
  (let* ((formats (select-tags-xlsx xlsx-file
				   "xl/styles.xml"
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
  (and (>= (length str) substring)
       (string= substring (subseq str 0 (length substring)))))

;; From Gwang-Jin Kim
(defun app-type (file)
  "Return the type of an .xlsx or .ods file."
  (let ((entries (list-entries file)))
    (flet ((extract-app-name (mode)
	     (let* ((file-is-ods (string= mode "ods"))
		    (xml  (if file-is-ods "meta.xml" "docProps/app.xml"))
		    (tags (if file-is-ods '(:meta :generator) '(:Application))))
	         (caddar (select-tags-xlsx file xml tags))))
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
	     "xlsx-microsoft"))))))) ;; works!

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
  (third (third table-cell)))

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
  (third (car (select-tags-xmlrep cell-tag '("v")))))

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
