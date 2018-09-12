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
  (let ((entry (zip:get-zip-file-entry name zip)))
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
  "Return tags matching the tags sequentially. 
   Unzip xlsx file (excel-path) and parses xml file, then sequentially select tags,
   while flattening results inbetween. Thus, output is a plain list of the targeted
   tag objects (list structures defined by the :xmls package)."
  (let ((content)) ;; initialize content symbol
    (with-open-xlsx (content xml excel-path)
	(collect-extract-exprs tags (list content)))))

(defun attr-val (attr tag)
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
	  collect (cons (attr-val "Id" rel)
			(attr-val "target" rel)))))

;; from Carlos Ungil
;; modified by Gwang-Jin Kim
(defun get-unique-strings (xlsx-file)
  "Return unique strings - necessary for parsing excel data."
  (let ((tags (select-tags-xlsx xlsx-file
				"xl/sharedStrings.xml"
				'(:si :t))))
    (loop for tag in tags
	  collect (cond ((equal (xmls:node-attrs tag) '(("space" "preserve")))
			 (xml:xmlrep-string-child tag))
			(t " "))))) ;; corrected by Gwang-Jin Kim 18-09-07

;; From Carlos Ungil
;; rewritten by Gwang-Jin Kim

(defun get-number-formats (xlsx-file)
  (let* ((formats (select-tags-xlsx xlsx-file
				   "xl/styles.xml"
				   '(:numFmts :numFmt)))
	 (format-codes (loop for fmt in formats
			     collect (cons (parse-integer
					    (attr-val "numFmtId" fmt))
					   (attr-val "formatCode" fmt))))
	 (styles (select-tags-xlsx xlsx-file
				   "xl/styles.xml"
				   '(:cellXfs :xf))))
    (loop for style in styles
	  collect (let ((fmt-id (parse-integer
				 (attr-val "numFmtId" style))))
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

(defun list-sheeets (file)
  "Retrieves the id and name of the worksheet in the .xlsx/.xlsm file."
  (let ((sheets (select-tags-xlsx file "xl/workbook.xml" '(:sheets :sheet))))
    (loop for sheet in sheets
	  with rels (get-relationships file)
	  for sheet-id   = (attr-val "sheetId" sheet)
	  for sheet-name = (attr-val "name"    sheet)
	  for sheet-rel  = (attr-val "id"      sheet)
	  collect (list (parse-integer sheet-id)
			sheet-name
			(cdr (assoc sheet-rel rels :test #'string=))))))

;; From Carlos Ungil
;; rewritten by Gwang-Jin Kim

(defun sheet-address (file sheet)
  "Return inner xml address of an excel sheet."
  (let* ((sheets (list-sheets file))
	 (entry-name (cond ((and (null sheet) (= 1 (length sheets)))
			    (caddr (car sheets)))
			   ((stringp sheet)
			    (caddr (find sheet
					 sheets
					 :key #'cadr
					 :test #'string=)))
			   ((numberp sheet)
			    (caddr (find sheet
					 sheets
					 :key #'car))))))
    (unless entry-name
      (error "specify one of the following sheet ids or names: ~{~&~{~S~^~5T~}~}"
	     (loop for (id name) in sheets
		   collect (list id name))))
    entry-name))
