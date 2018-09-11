;;;; cl-xlsx.asd

(asdf:defsystem #:cl-xlsx
  :name "xl"
  :description "Basic reader for .xlsx and .ods files"
  :author "Gwang-Jin Kim <gwang.jin.kim.phd@gmail.com>"
  :license "MIT"
  :serial t
  :components ((:file "package")
               (:file "cl-xlsx"))
  :depends-on (:zip :flexi-streams :xmls :babel))

