resolvers += 
  "Apache Releases" at "https://repository.apache.org/content/repositories/releases"

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi-ooxml" % "3.17",
  "org.apache.xmlbeans" % "xmlbeans" % "3.0.0",
  "xerces" % "xercesImpl" % "2.12.0"
)
