(ns clj-xlsx.xlsx.test
  (:require [clojure.test :refer [deftest is testing]]
            [clojure.string :as str]
            [clj-xlsx.share.schema.test :as test-sch]
            [clj-xlsx.core :as xlsx]))


(defn- cell-style-match?
  [_ _] true)

(defn- cell-match?
  [{:keys [v f s]}
   {c-v :v
    c-f :f
    c-s :s}]
  (and (= c-f f)
       (= c-v v)
       (cell-style-match? s c-s)))

(defn data-match?
  [d c-d]
  (mapv (fn [{:keys [name col-styles rows]}
             {c-name :name c-col-style :col-styles c-rows :rows}]
          (and (= name c-name)
               (every? true?
                       (mapv cell-match? rows c-rows))))
        d c-d))


(deftest rw-xlsx-test
  (->> test-sch/test-data
       (mapv (fn [[case-id data]]
               (let [[s n] (->> case-id
                                name
                                ((fn [s] (str/split s #"-"))))
                     case-s    [s (Integer/parseInt n)]
                     file-name (apply format "%s-%02d.xlsx" case-s)
                     label     (apply format "Testing writing/reading xlsx %s %02d" case-s)]
                 (testing label
                   (let [_ (->> data
                                xlsx/->excel-byte-array
                                (xlsx/write-file file-name))
                         r (xlsx/read-workbook file-name)]
                     (is (data-match? data r)))))))))