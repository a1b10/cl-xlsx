(ql:quickload :cl-xlsx)
(in-package :cl-xlsx)
;; the package is so unintuitive
;; and helper functions for development needed



;; write a function to

(zip:with-zipfile (fin (truename "/home/josephus/Dropbox/cl/sources/cl-xlsx/test.xlsx"))
  (zip:zipfile-entries fin))

(defparameter *fp* "/home/josephus/Dropbox/cl/sources/cl-xlsx/test.xlsx")

(defun list-entries (file)
  "Internal use, gets entries inside of ZIP/XLSX files."
  (zip:with-zipfile (zip file)
    (let ((entries (zip:zipfile-entries zip)))
      (when entries
	(loop for k being the hash-keys of entries
	   collect k)))))

(list-sheets *fp*)

(read-xlsx *fp*)
