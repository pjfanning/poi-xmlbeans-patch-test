# poi-poi-xmlbeans-patch-test

## Running the sample code

- [Install SBT](http://www.scala-sbt.org/)
- `sbt "run [filename]"`


The test code will extract the text from the spreadsheet. If you don't supply a filename, it will read the sample.xlsx.

sample.xlsx has some Greek characters that are [Unicode Surrogates](https://en.wikipedia.org/wiki/Universal_Character_Set_characters#Surrogates)
and these characters are replaced by '?' characters when we write the workbook out. com.github.pjfanning:xmlbeans jar fixes this issue.

If you want to check the behaviour without the org.apache.xmlbeans jar, modify the build.sbt and remove the exclude and omit the
com.github.pjfanning version of the xmlbeans jar.
