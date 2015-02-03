(defproject clj-excel "0.0.1"
  :description "Excel bindings for Clojure, based on Apache POI."
  :dependencies [[org.clojure/clojure "1.6.0"]
                 [org.apache.poi/poi "3.11"]
                 [org.apache.poi/poi-ooxml "3.11"]]

  ;; lein with-profile dev cloverage [cloverage-opts]
  :profiles {:dev {:source-paths ["dev"]
                   :resource-paths ["test-resources"]
                   :plugins [[lein-cloverage "1.0.2"]]
                   :dependencies [[org.clojure/tools.namespace "0.2.4"]]
                   :global-vars {*print-length* 20}}
             :test {:resource-paths ["test-resources"]}}
  :global-vars {*warn-on-reflection* true})
