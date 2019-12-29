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
  :depends-on (:cxml :zip :babel :xpath :fxml :parse-number :local-time))
