(defpackage :cl-xlsx
  (:use :common-lisp)
  (:export #:read-xlsx
	   #:sheet-names))


(in-package #:cl-xlsx)


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

(defun flatten (l &key (acc '()))
  "Flatten a list - to avoid dependency."
  (cond ((null l) acc)
	((atom (car l)) (flatten (cdr l) :acc (append acc (list (car l)))))
	(t (flatten (cdr l) :acc (append acc (flatten (car l)))))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; cxml
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun source-entry (xml xlsx)
  "Get content xml file inside xlsx."
  (zip:with-zipfile (zip xlsx)
    (let ((entry (zip:get-zipfile-entry xml zip)))
      (when entry
	(cxml:make-source (babel:octets-to-string (zip:zipfile-entry-contents entry)))))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; sax parsing
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun sax-attributes (sax)
  (second sax))

(defun sax-attribute (sax attribute)
  "Return value of attribute in sax."
  (let ((attributes (sax-attributes sax)))
    (car (last (assoc attribute attributes :key #'car :test #'string=)))))

;;; extractor functions for sax-parsed cell

(defun sax-cell-value (sax)
  "Return value of sax-parsed cell tag (in child tag :v)."
  (car (last (car (last sax)))))

(defun sax-cell-type (sax)
  "Return type of sax-parsed excel cell tag (attribute 't')."
  (cadr (assoc "t" (sax-attributes sax) :test #'equal)))

(defun sax-cell-pos (sax)
  "Return position information of sax-parsed excel cell tag (attribute 'r')."
  (cadr (assoc "r" (sax-attributes sax) :test #'equal)))

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

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; read xlsx
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun inner-files (xlsx)
  "List all innter addresses in xlsx."
  (zip:with-zipfile (inzip xlsx)
    (let ((entries (zip:zipfile-entries inzip)))
      (when entries
	(loop for k being the hash-keys of entries
	   collect k)))))

(defun sheets (xlsx)
  "Return sheet informations as list of lists (sheet-name sheet-number sheet-address)."
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

#| actually:
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

(defun sheet-names-xlsx (xlsx)
  "List sheet names in xlsx file."
  (mapcar #'car (sheets xlsx)))

(defun sheet-address (sheet xlsx)
  "Return sheet address inside xlsx file when sheet name or index given as input."
  (typecase sheet
    (string (caddr (assoc sheet (sheets xlsx) :test #'string=)))
    (integer (cadr (assoc sheet (mapcar #'cdr (sheets xlsx)))))))

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



(defun parse-xlsx-sheet (sheet xlsx)
  "Return parsed content for a given sheet in a xlsx-fpath."
  (klacks:with-open-source (s (source-entry (sheet-address sheet xlsx) xlsx))
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

(defun parse-xlsx (xlsx)
  "Parse every sheet of xlsx and return as alist (sheet-name sheet-content-as-list)."
  (let ((sheet-names (sheet-names-xlsx xlsx)))
    (mapcar #'(lambda (sheet)
		(cons sheet (parse-xlsx-sheet sheet xlsx)))
	    sheet-names)))



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

(defun sheet-names-ods (ods)
  "Return sheet names."
 (klacks:with-open-source (s (source-entry "content.xml" ods))
   (loop for key = (klacks:peek s)
	for consumed = nil
      while key
      nconcing (case key
		 (:start-element
		  (when (equal (klacks:current-qname s) "table:table")
		    (setf consumed t)
		    (let ((sax (klacks:serialize-element s (cxml-xmls:make-xmls-builder))))
		      (list (sax-attribute sax "name"))))))
       do (unless consumed
	    (klacks:consume s))))) ; works!

(defun number-columns-repeated-cell-p (table-cell-sax)
  "Tester whether 'number-columns-repeated' table-cell-sax or not."
  (equal (car (caaadr table-cell-sax)) "number-columns-repeated"))

;; requires flatten above
(defun all-nil-row-p (l)
  (every #'null (flatten l)))

(defun parse-ods (ods)
  "Return parsed content for xlsx-fpath."
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



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; read xlsx/ods recognizing automatically
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

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

(defun sheet-names (xlsx)
  "Return sheet-names of xlsx."
  (let ((type (app-type xlsx)))
    (cond ((or (string= type "xlsx-microsoft")
	       (string= type "xlsx-libreoffice"))
	   (sheet-names-xlsx xlsx))
	  ((string= type "ods-libreoffice")
	   (sheet-names-ods xlsx))
	  (t nil))))

(defun read-xlsx (xlsx)
  "Read xlsx and ods file with all sheets into a list of list of lists - 
   recognizing automatically type of file (xlsx or ods/odt)."
  (let ((type (app-type xlsx)))
    (cond ((or (string= type "xlsx-microsoft")
	       (string= type "xlsx-libreoffice"))
	   (parse-xlsx xlsx))
	  ((string= type "ods-libreoffice")
	   (loop for sheet-name in (sheet-names-ods xlsx)
		 for sheet-content in (parse-ods xlsx)
		 collect (cons sheet-name sheet-content)))
	  (t nil))))

;; (defun read-xlsx (xlsx &key (sheet nil))
;;   "Read xlsx and ods file with all sheets into a list of list of lists - 
;;    recognizing automatically type of file (xlsx or ods/odt)."
;;   (let ((type (app-type xlsx)))
;;     (if (null sheet)
;; 	(cond ((or (string= type "xlsx-microsoft")
;; 		   (string= type "xlsx-libreoffice"))
;; 	       (cond ((null sheet) (parse-xlsx xlsx))
;; 		     ((atom sheet) (parse-xlsx-sheet xlsx))
;; 		     ((consp sheet) (mapcar 
;; 		   )
;; 	      ((string= type "ods-libreoffice")
;; 	       (loop for sheet-name in (sheet-names-ods xlsx)
;; 		     for sheet-content in (parse-ods xlsx)
;; 		collect (cons sheet-name sheet-content)))
;; 	      (t nil))
;; 	(cond (())))))


