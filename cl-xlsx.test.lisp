(ql:quickload :cl-xlsx)

(cl-xlsx::parse-xlsx "~/test/xlsx-files/test.xlsx")
(cl-xlsx::sheets "~/test/xlsx-files/test.xlsx")
(cl-xlsx::sheets "~/test/xlsx-files/test.ods")

(cl-xlsx::read-xlsx "~/test/xlsx-files/test.xlsx")
(cl-xlsx::read-xlsx "~/Dropbox/amit_scripts/cnts_avgs_DES2nrm.xlsx") ; that works
;;; maybe because unique strings function is overwritten
(defparameter *x* (cl-xlsx::read-xlsx "~/Dropbox/amit_scripts/cnts_avgs_DES2nrm.xlsx"))
(length (elt *x* 0))
(length (elt (elt *x* 0) 0))

(defun head (table &key (n 5))
  (subseq table 0 n))

(head (elt *x* 0)) ; works
  
(defun tail (table &key (n 5))
  (let ((ln (length table)))
    (subseq table (- ln n) ln))) 

(tail (elt *x* 0)) ; works as intended


(cl-xlsx::
