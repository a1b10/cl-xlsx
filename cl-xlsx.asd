;;;; cl-xlsx.asd

(asdf:defsystem #:cl-xlsx
  :name "cl-xlsx"
  :version "1.0"
  ;; :description "Basic reader for .xlsx and .ods files using streams"
  :author "Gwang-Jin Kim <gwang.jin.kim.phd@gmail.com>"
  :license "MIT"
  :serial t
  :components ((:file "package")
               (:file "cl-xlsx"))
  :description "Read LibreOffice ODS files and LibreOffice and Microsoft XLSX files using Common Lisp"
  :depends-on (:cxml :zip :babel :xpath :fxml :fxml/cxml :fxml/stp :fxml/xpath :parse-number :local-time)
  :perform (load-op :after (op c)
             (format t "NOTICE: cl-xlsx will not be further developed. Please use cl-excel instead.~%")
             (format t "cl-excel allows you to read & write your tables from/into Excel sheets! ~%")
             (format t "For installation, see the instructions in https://github.com/gwangjinkim/cl-excel . ~%"))
