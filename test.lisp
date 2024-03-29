(ql:quickload :cl-xlsx)

(defparameter *fp* "~/Dropbox/cl/sources/cl-xlsx/test.xlsx")
(cl-xlsx:read-xlsx *fp*) ;; works!

(defparameter *f* "Dropbox/Rprojects/xlsx2dfs/test_one_df.xlsx")
(cl-xlsx:read-xlsx *f*)


(defparameter *f1* "Dropbox/amit_scripts/org.Ct.eg.db/test_data.xlsx")
(cl-xlsx:read-xlsx *f1*) ;; works!

(defparameter *f2* "Dropbox/amit_scripts/alaAmit/DE_dKOvsWT_down_0.01.xlsx")
(cl-xlsx:read-xlsx *f2*)

;; Couldn't find child tag V in ((c
;;                                . http://schemas.openxmlformats.org/spreadsheetml/2006/main)
;;                               ((r A1)))


(defparameter *f3* "Dropbox/amit_scripts/clustering/kmeansClusteringCellsJelena.xlsx")
(cl-xlsx:read-xlsx *f3*) ;; works

(defparameter *f4* "Dropbox/amit_scripts/RawCounts.xlsx")
(cl-xlsx:read-xlsx *f4*)

;; Dropbox/amit_scripts/RawCounts.xlsx
;; Control stack exhausted (no more space for function call frames).
;; This is probably due to heavily nested or infinitely recursive function
;; calls, or a tail call that SBCL cannot or has not optimized away.

;; PROCEED WITH CAUTION.

;; > 45000 rows - maybe too many recursive calls!



;;;;;;;;;;;;;;;;;;;;;;;;;;
;; testing reading data files for ujo
;;;;;;;;;;;;;;;;;;;;;;;;;;

(ql:quickload :cl-xlsx)
(defparameter *bbc* "/home/josephus/Downloads/datasamples/BBC_2018Q1.xlsx")
(defparameter *u1* "/home/josephus/Downloads/datasamples/userdata1.xlsx")
(defparameter *u2* "/home/josephus/Downloads/datasamples/userdata2.xlsx")

(defparameter *tbbc* (cl-xlsx:read-xlsx *bbc*))
(elt *tbbc* 0) ;; unconnected!
(defparameter *u1* (cl-xlsx:read-xlsx *u1*))
(defparameter *u2* (cl-xlsx:read-xlsx *u2*))
*u2* ;; works
