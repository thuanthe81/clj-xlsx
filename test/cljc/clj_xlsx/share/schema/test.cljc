(ns clj-xlsx.share.schema.test
  (:require [clojure.test :refer [deftest is testing]]
            [clj-xlsx.share.schema :as sch]
            [schema.core :as s]
            [clojure.string :as str]))

;;(def date-ts 61538659200000) ;; "2020-01-01T00:00:00Z
(def date-ts (java.util.Date. (System/currentTimeMillis) #_(java.util.Date. 2020 01 01 00 00 00)))

(def test-data
  {:case-1 [{:name "Pricing"
             :rows [[{:v "Product"
                      :s {:font                 {:name "Arial"
                                                 :bold true
                                                 :size 20}
                          :horizontal-alignment :center
                          :vertical-alignment   :center
                          :background-color     "#D5DBD6"}}
                     nil
                     {:v "Price"
                      :s {:font                 {:name "Arial"
                                                 :bold true
                                                 :size 20}
                          :horizontal-alignment :center
                          :vertical-alignment   :center}}]]}]
   :case-2 [{:name       "Formula"
             :col-styles [nil {:width 30} {:width :auto}]
             :rows       [[{:v 15
                            :s {:font {:name "Arial"
                                       :size 12}}}
                           {:v 54
                            :s {:font {:name "Arial"
                                       :size 12}}}
                           {:v 69
                            :f "SUM(A1:A2)"
                            :s {:font {:name   "Arial"
                                       :size   16
                                       :color  "#FF0000"
                                       :italic true}}}]]}]
   :case-3 [{:name       "Formatting"
             :col-styles [{:width :auto}]
             :rows       [[{:v   date-ts
                            :s   {:font {:name "Arial"
                                         :size 12}
                                  :fmt {:n 0xe}}}]]}]
   :case-4 [{:name       "Empty cells"
             :rows      [[{:v "row 1"} {:v 1}]
                         [{:v "row 2"} {:v 2}]
                         nil
                         []
                         nil
                         []
                         [{:v "total"} {:f "SUM(B1:B2)"}]]}]})

(deftest schema-test
  (->> test-data
       (mapv (fn [[case-id data]]
               (let [label (->> case-id
                                name
                                ((fn [s] (str/split s #"-")))
                                (apply format "Testing schema %s %s"))]
                 (testing label
                   (is (s/validate sch/sheets data))))))))