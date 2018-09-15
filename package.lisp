;;;; package.lisp

(defpackage #:cl-xlsx
  (:use #:cl)
  (:export #:list-entries
	   #:get-entry
	   #:with-open-xlsx
	   #:once-flatten
	   #:extract-sub-tags
	   #:collect-extract-exprs
	   #:select-tags-xlsx
	   #:select-tags-xmlrep
	   #:attr-val
	   #:get-relationships
	   #:get-unique-strings
	   #:get-number-formats
	   #:column-and-row
	   #:excel-date
	   #:list-sheets
	   #:sheet-address
	   #:begins-with?
	   #:app-type
	   #:process-table-cell
	   #:process-table-row
	   #:process-table-rows-ods
	   #:read-ods))
