;;;; cl-xlsx.lisp

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


;;;-----------------------------------------------------------------------------
;;; cxml
;;;-----------------------------------------------------------------------------

(defun source-entry (xml xlsx)
  "Get content xml file inside xlsx."
  (zip:with-zipfile (zip xlsx)
    (let ((entry (zip:get-zipfile-entry xml zip)))
      (when entry
        (cxml:make-source (babel:octets-to-string (zip:zipfile-entry-contents entry)))))))


;;;-----------------------------------------------------------------------------
;;; date/time stuff
;;;-----------------------------------------------------------------------------

(defun excel-date-serial-number-to-timestamp (dsn &optional (timezone local-time:*default-timezone*))
  (local-time:adjust-timestamp
      (local-time:encode-timestamp 0 1 0 0 1 1 1900 :timezone timezone)
    (offset :day (1- (if (< dsn 60)
                         dsn
                         (1- dsn))))))

(defun decode-excel-date-serial-number (dsn)
  (local-time:with-decoded-timestamp
      (:year year :month month :day day :timezone local-time:+utc-zone+)
      (excel-date-serial-number-to-timestamp dsn)
    (values year month day)))

(defun print-iso-date (lt-date &optional stream (timezone local-time:*default-timezone*))
  (local-time:format-timestring stream lt-date
                                :format local-time:+iso-8601-date-format+
                                :timezone timezone))

;;;-----------------------------------------------------------------------------
;;; xpath stuff
;;;-----------------------------------------------------------------------------


(eval-when (:compile-toplevel :load-toplevel :execute)
  (defparameter *package-relationships-namespace*
    "http://schemas.openxmlformats.org/package/2006/relationships")

  (defparameter *default-package-relationships-namespace*
    `(("" ,*package-relationships-namespace*)))

  (defparameter *office-document-relationships-namespace*
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

  (defparameter *default-xlsx-namespaces*
    `(("" "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      ("mc" "http://schemas.openxmlformats.org/markup-compatibility/2006")
      ("r" ,*office-document-relationships-namespace*)
      ("x15" "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
      ("xr" "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
      ("xr2" "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
      ("xr6" "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
      ("xr10" "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"))))

(defmacro with-package-relationships-namespaces ((&optional extra-namespaces) &body body)
  `(xpath:with-namespaces ,(append *default-package-relationships-namespace*
                                   extra-namespaces)
     ,@body))

(defmacro with-xlsx-namespaces ((&optional extra-namespaces) &body body)
  `(xpath:with-namespaces ,(append *default-xlsx-namespaces* extra-namespaces)
     ,@body))

(defun guess-date-format-p (fmt-string)
  (and (find #\D fmt-string :test 'char-equal)
       (find #\M fmt-string :test 'char-equal)
       (find #\Y fmt-string :test 'char-equal)))

(defun find-user-defined-style (styles style-num)
  (let ((user-defined-style (cdr (assoc (elt (first styles) style-num)
                                        (second styles)))))
    (when user-defined-style
      (cond ((guess-date-format-p user-defined-style)
             :date)
            (t :general)))))

(defun get-built-in-style-type (style-num)
  (cond ((zerop style-num)
         :general)
        ((or (<= 1 style-num 4)
             (<= 37 style-num 40)
             (= style-num 48))
         :numeric)
        ((<= 9 style-num 10)
         :percent)
        ((<= 14 style-num 17)
         :date)
        ((or (<= 18 style-num 21)
             (<= 45 style-num 47))
         :time)
        ((= 22 style-num)
         :date-time)
        (t :general)))

(defun excel-col-index (col)
  (flet ((char-num (char)
           (1+ (- (char-code (char-upcase char)) (char-code #\A)))))
    (1- (reduce (lambda (x y)
                  (+ (* x 26) (char-num y))) col :initial-value 0))))

(defun process-cell-node (node unique-strings styles current-col)
  "Return value of cell node. Checks type 't' and if string 's', looks
   up from unique-strings the right string.  If numeric 'n', then
   parses it using parse-number:parse-number, Otherwise return the
   cell value directly."
  (declare (optimize (debug 3)))
  (with-xlsx-namespaces ()
    (let ((element (fxml.stp:first-child node)))
      (fxml.stp:with-attributes ((pos "r")
                                 (cell-style "s")
                                 (cell-type "t"))
          element
        (let ((val (let ((value-node (xpath:first-node (xpath:evaluate "v/text()" element))))
                     (cond ((equalp cell-type "s")
                            (elt unique-strings (parse-integer (xpath:string-value value-node))))
                           ((equalp cell-type "n")
                            (let ((cell-value
                                   (parse-number:parse-number (xpath:string-value value-node))))
                              (if cell-style
                                  (let ((cell-style-num (parse-integer cell-style)))
                                    (if cell-style-num
                                        (let ((cell-style
                                               (or (find-user-defined-style styles cell-style-num)
                                                   (get-built-in-style-type cell-style-num))))
                                          (case cell-style
                                            ((:date :date-time)
                                             (excel-date-serial-number-to-timestamp
                                              cell-value))
                                            (t cell-value)))
                                        cell-value))
                                  cell-value)))
                           (t
                            (when value-node
                              (xpath:string-value value-node)))))))
          (let* ((pos-first-num (position-if #'digit-char-p pos))
                 (col (subseq pos 0 pos-first-num))
                 (col-index (excel-col-index col)))
            (if (> col-index current-col)
                (append (make-list (- col-index current-col))
                        (list val))
                (list val))))))))

;;;-----------------------------------------------------------------------------
;;; sax parsing
;;;-----------------------------------------------------------------------------

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
          ((equalp val-type "n") (parse-number:parse-number (sax-cell-value sax)))
          (t (sax-cell-value sax)))))

;;;-----------------------------------------------------------------------------
;;; read xlsx
;;;-----------------------------------------------------------------------------

(defun inner-files (xlsx)
  "List all innter addresses in xlsx."
  (zip:with-zipfile (inzip xlsx)
    (let ((entries (zip:zipfile-entries inzip)))
      (when entries
        (loop :for k :being :the :hash-keys :of entries
              :collect k)))))

(defun get-relationships (xlsx)
  "Returns sheet relationships as list of lists (...)."
  (let ((rels (fxml:parse (source-entry-stream "xl/_rels/workbook.xml.rels" xlsx)
                              (fxml.stp:make-builder))))
    (with-package-relationships-namespaces ()
      (xpath:map-node-set->list
       (lambda (rel-node)
         (fxml.stp:with-attributes ((id "Id")
                                    (type "Type")
                                    (target "Target"))
             rel-node
           (list id type target)))
       (xpath:evaluate "/Relationships/Relationship" rels)))))

(defun get-sheets (xlsx)
  "Return sheet informations as list of lists (sheet-name sheet-number sheet-id sheet-address)."
  (let ((relationships (get-relationships xlsx)))
    (let ((workbook (fxml:parse (source-entry-stream "xl/workbook.xml" xlsx)
                                (fxml.stp:make-builder))))
      (with-xlsx-namespaces ()
        (xpath:map-node-set->list
         (lambda (sheet-node)
           (fxml.stp:with-attributes ((r-id "r:id" *office-document-relationships-namespace*)
                                      (sheet-id "sheetId")
                                      (name "name"))
               sheet-node
             (let ((related-item (find r-id relationships :key #'car :test 'equal)))
               (list name sheet-id r-id (elt related-item 2)))))
         (xpath:evaluate "/workbook/sheets/sheet" workbook))))))

(defun sheets (xlsx)
  "Return sheet informations as list of lists (sheet-name sheet-number sheet-address)."
  (klacks:with-open-source (src (source-entry "xl/workbook.xml" xlsx))
    (loop :for key = (klacks:peek src)
	  :for consumed = nil
          :while key
          :nconcing (case key
		      (:start-element
		       (let ((tag-name (klacks:current-qname src)))
			 (cond ((equal tag-name "sheet")
				(setf consumed t)
				(list (let* ((sax (klacks:serialize-element src (cxml-xmls:make-xmls-builder)))
					     (attributes (cadr sax)))
					(list (cadr (assoc "name" attributes :test #'equal))
					      (parse-integer (cadr (assoc "sheetId" attributes :test #'equal)))
					      (concatenate 'string "xl/worksheets/sheet"
							   (cadr (assoc "sheetId" attributes :test #'equal))
							   ".xml")))))
			       (t nil))))
		      (otherwise nil))
	  :do (unless consumed
		(klacks:consume src)))))

;; (defun sheets (xlsx)
;;   "Return sheet informations as list of lists (name sheet-number sheetaddress."
;;   (let ((sheet-addresses (remove-if-not #'(lambda (x) (and (starts-with-p x "xl/worksheets/")
;;                                                            (ends-with-p x ".xml")))
;;                                         (inner-files xlsx))))
;;     (klacks:with-open-source (src (source-entry "xl/workbook.xml" xlsx))
;;       (loop for key = (klacks:peek src)
;;             while key
;;             nconc (case key
;;                     (:start-element
;;                      (if (equal (klacks:current-qname src) "sheet")
;;                          (list (let* ((sax (klacks:serialize-element src (cxml-xmls:make-xmls-builder)))
;;                                       (attributes (cdadr sax)))
;;                                  (list (cadr (assoc "name" attributes :test #'equal))
;;                                        (parse-integer (cadr (assoc "sheetId" attributes :test #'equal)))
;;                                        (pop sheet-addresses))))
;;                          nil))
;;                     (otherwise nil))
;;             do (klacks:consume src)))))
;; works

(defun sheet-names-xlsx (xlsx)
  "List sheet names in xlsx file."
  (mapcar #'car (sheets xlsx)))

(defun sheet-address (sheet xlsx)
  "Return sheet ID when sheet name or index given as input. Note that
this is not the same as the user-visible sheet number and should not
be used as an offset into the list of sheets. Use get-sheets to get
more detailed information on sheets for that."
  (typecase sheet
    (string (caddr (assoc sheet (sheets xlsx) :test #'string=)))
    (integer (cadr (assoc sheet (mapcar #'cdr (sheets xlsx)))))))

(defun source-entry-stream (xml xlsx)
  "Get content xml file inside xlsx."
  (zip:with-zipfile (zip xlsx)
    (let ((entry (zip:get-zipfile-entry xml zip)))
      (when entry
        (zip:zipfile-entry-contents entry)))))

(defun get-unique-strings (xlsx)
  (let ((doc (fxml:parse (source-entry-stream "xl/sharedStrings.xml" xlsx)
                         (fxml.stp:make-builder))))
    (with-xlsx-namespaces ()
      (xpath:map-node-set->list
       (lambda (x)
         (let ((rich-text-runs (xpath:evaluate "r/t/text()" x)))
           (if (not (xpath:node-set-empty-p rich-text-runs))
               (progn
                 (apply #'concatenate
                        'string
                        (xpath:map-node-set->list #'xpath:string-value rich-text-runs)))
               (let ((text (xpath:string-value (xpath:evaluate "t/text()" x))))
                 (if text
                     text
                     )))))
       (xpath:evaluate "/sst/si" doc)))))

(defun get-doc-number-formats (doc)
  (with-xlsx-namespaces ()
    (xpath:map-node-set->list
     (lambda (x)
       (fxml.stp:with-attributes ((format-code "formatCode")
                                  (number-format-id "numFmtId"))
           x
         (cons (parse-integer number-format-id) format-code)))
     (xpath:evaluate "/styleSheet/numFmts/numFmt" doc))))

(defun get-doc-cell-formats (doc)
  (with-xlsx-namespaces ()
    (xpath:map-node-set->list
     (lambda (x)
       (fxml.stp:with-attributes ((number-format-id "numFmtId"))
           x
         (parse-integer number-format-id)))
     (xpath:evaluate "/styleSheet/cellXfs/xf" doc))))

(defun get-styles (xlsx)
  (let ((doc (fxml:parse (source-entry-stream "xl/styles.xml" xlsx)
                         (fxml.stp:make-builder))))
    (list (get-doc-cell-formats doc)
          (get-doc-number-formats doc))))

(defun parse-xlsx-sheet (sheet-address xlsx &key unique-strings styles keep-empty-rows-as-nil)
  "Return parsed content for a given sheet in a xlsx-fpath."
  (klacks:with-open-source (s (source-entry (concatenate 'string "xl/" sheet-address) xlsx))
    (let ((unique-strings (or unique-strings
                              (get-unique-strings xlsx))))
      (loop :for key = (klacks:peek s)
         :with row-counter = 1
         :while key
         :nconcing (case key
                     (:start-element
                      (let ((tag-name (klacks:current-qname s)))
                        (cond ((equal tag-name "row")
                               (let ((row-num (parse-number:parse-number (klacks:get-attribute s "r"))))
                                 (loop :for key = (klacks:peek s)
                                    :for consumed = nil
                                    :with i = 0
                                    :while key
                                    :nconcing (case key
                                                (:start-element
                                                 (cond ((equal (klacks:current-qname s) "c")
                                                        (setf consumed t)
                                                        (let ((cols
                                                               (progn
                                                                 (process-cell-node
                                                                  (klacks:serialize-element
                                                                   s
                                                                   (fxml.stp:make-builder))
                                                                  unique-strings
                                                                  styles
                                                                  i))))
                                                          (incf i (length cols))
                                                          cols))
                                                       (t
                                                        (setf consumed nil)
                                                        nil)))
                                                (:end-element
                                                 (when (equal (klacks:current-qname s) "row")
                                                   (let ((pad-rows (- row-num row-counter)))
                                                     (incf row-counter (1+ pad-rows))
                                                     (return (append (when keep-empty-rows-as-nil
                                                                       (make-list pad-rows))
                                                                     (list res)))))))
                                    :into res
                                    :do (unless consumed
                                          (klacks:consume s)))))
                              (t nil))))
                     (otherwise nil))
         :do (klacks:consume s))))) ;; works!

(defun parse-xlsx (xlsx &key keep-empty-rows-as-nil)
  "Parse every sheet of xlsx and return as alist (sheet-name sheet-content-as-list)."
  (let ((sheets (get-sheets xlsx))
        (unique-strings (get-unique-strings xlsx))
        (styles (get-styles xlsx)))
    (mapcar #'(lambda (sheet)
                (destructuring-bind (sheet-name sheet-number sheet-id sheet-address)
                    sheet
                  (declare (ignore sheet-number sheet-id))
                  (cons sheet-name (parse-xlsx-sheet sheet-address xlsx
                                                     :unique-strings unique-strings
                                                     :styles styles
                                                     :keep-empty-rows-as-nil keep-empty-rows-as-nil))))
            sheets)))


;;;-----------------------------------------------------------------------------
;;; read ods
;;;-----------------------------------------------------------------------------

(defun parse-ods-cell (sax)
  "Return string or number according to type of ods cell sax."
  (let ((type (car (last (first (second sax)))))
        (value (car (last (third sax)))))
    (cond ((equalp type "float") (parse-number:parse-number value))
          ;; since we return the same for "string" and the default
          ;; case, no need to check "string" here.
          ;;
          ;; ((equalp type "string") value)
          (t value))))

(defun sheet-names-ods (ods)
  "Return sheet names."
  (klacks:with-open-source (s (source-entry "content.xml" ods))
    (loop :for key = (klacks:peek s)
          :for consumed = nil
          :while key
          :nconcing (case key
                     (:start-element
                      (when (equal (klacks:current-qname s) "table:table")
                        (setf consumed t)
                        (let ((sax (klacks:serialize-element s (cxml-xmls:make-xmls-builder))))
                          (list (sax-attribute sax "name"))))))
          :do (unless consumed
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
    (loop :for key = (klacks:peek s)
          :while key
          :nconcing (case key
                     (:start-element
                      (cond ((equal (klacks:current-qname s) "table:table")
                             (loop :for key1 = (klacks:peek s)
                                   :while key1
                                   :nconcing (case key1
                                              (:start-element
                                               (cond ((equal (klacks:current-qname s) "table:table-row")
                                                      (loop :for key2 = (klacks:peek s)
                                                            :for consumed = nil
                                                            :while key2
                                                            :nconcing (case key2
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
                                                            :into inner-res
                                                            :do (unless consumed
                                                                 (klacks:consume s))))
                                                     (t nil)))
                                              (:end-element
                                               (if (equal (klacks:current-qname s) "table:table")
                                                   (return (list res))))
                                              (otherwise nil))
                                   :into res
                                   :do (klacks:consume s)))
                            (t nil)))
                     (otherwise nil))
          :do (klacks:consume s))))


;;;-----------------------------------------------------------------------------
;;; read xlsx/ods recognizing automatically
;;;-----------------------------------------------------------------------------

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
          ((or (starts-with-p (app-name xlsx) "Microsoft Excel")
               (starts-with-p (app-name xlsx) "Microsoft Macintosh Excel")
			   (starts-with-p (app-name xlsx) "WPS 表格")
			   (starts-with-p (app-name xlsx) "Apache POI"))
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

(defun read-xlsx (xlsx &key keep-empty-rows-as-nil)
  "Read xlsx and ods file with all sheets into a list of list of lists -
   recognizing automatically type of file (xlsx or ods/odt)."
  (let ((type (app-type xlsx)))
    (cond ((or (string= type "xlsx-microsoft")
               (string= type "xlsx-libreoffice"))
           (parse-xlsx xlsx :keep-empty-rows-as-nil keep-empty-rows-as-nil))
          ((string= type "ods-libreoffice")
           (loop :for sheet-name :in (sheet-names-ods xlsx)
                 :for sheet-content :in (parse-ods xlsx)
                 :collect (cons sheet-name sheet-content)))
          (t nil))))

;; (defun read-xlsx (xlsx &key (sheet nil))
;;   "Read xlsx and ods file with all sheets into a list of list of lists -
;;    recognizing automatically type of file (xlsx or ods/odt)."
;;   (let ((type (app-type xlsx)))
;;     (if (null sheet)
;;      (cond ((or (string= type "xlsx-microsoft")
;;                 (string= type "xlsx-libreoffice"))
;;             (cond ((null sheet) (parse-xlsx xlsx))
;;                   ((atom sheet) (parse-xlsx-sheet xlsx))
;;                   ((consp sheet) (mapcar
;;                 )
;;            ((string= type "ods-libreoffice")
;;             (loop for sheet-name in (sheet-names-ods xlsx)
;;                   for sheet-content in (parse-ods xlsx)
;;              collect (cons sheet-name sheet-content)))
;;            (t nil))
;;      (cond (())))))
