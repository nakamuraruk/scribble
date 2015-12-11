(ns scribble.core
  (:require [clojure.java.io :as io])
  (:import 
    (org.apache.poi.ss.usermodel WorkbookFactory)))

(defn load-workbook [efile]
  (WorkbookFactory/create (io/file efile)))

(defn get-sheet [sheetname] 
  (.getSheet sheetname))

(defn get-row [sheet rownum]
  (.getRow sheet rownum))
