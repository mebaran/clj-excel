(defproject grafter/clj-excel "0.0.10-SNAPSHOT"
  :description "Excel bindings for Clojure, based on Apache POI."
  :dependencies [[org.clojure/clojure "1.8.0"]
                 [org.apache.poi/poi "3.9"]
                 [org.apache.poi/poi-ooxml "3.9"]]
  :profiles {:dev {:source-paths ["dev"]
                   :resource-paths ["test-resources"]
                   :plugins [[lein-cloverage "1.0.2"]]
                   :dependencies [[org.clojure/tools.namespace "0.2.11"]]
                   :global-vars {*print-length* 20}}
             :test {:resource-paths ["test-resources"]}}
  :global-vars {*warn-on-reflection* true}
  :deploy-repositories  [["releases" :clojars]
                         ["snapshots" :clojars]])
