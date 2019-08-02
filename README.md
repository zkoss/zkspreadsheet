# ZK Spreadsheet

ZK Spreadsheet is an open source embeddable web-based online spreadsheet that delivers the rich functionality of Excel within browsers using pure Java. With embeddable Excel functionality, developers are able to create interactive, collaborative, and dynamic enterprise applications easily.

> Newer. Better. Faster.
> Check out [Keikai](https://keikai.io) - 2019 brand-new Web Spreadsheet from ZK Team.


# Binary Release Build Process
## ZSS
`bin/build_binary.py`

## zssapp
### official
`mvn package`

### Eval
```
mvn -P eval
mvn package
```

# Naming Convention for abbreviation names
1) Industry standard names, e.g., XML, HTML and URI:
	We use all uppercase letters as other projects do.
	For example, XMLs for the XML utilities

2) ZK's names, e.g., ZK and ZUL:
	We use uppercase for the first letter.
	For example, ZkFns for the ZK EL functions, UiFactory for the UI factory

---
zss
	The ZK Spreadsheet component.
	It includes ZK Spreadsheet component and associated data model APIs that you can
	use to manipulate your spreadsheet documents.

---
zssapp
	The ZK Spreadsheet Live.
	A Web based Excel(R) like application implemented with ZK Spreadsheet component
	and ZK Ajax framework(http://www.zkoss.org).

---
jsdoc
	The document of ZK spreadsheet JavaScript code(work with ZK Client Engine)
	
---
zpoi
	A library based on a part of the Apache POI spreadsheet library(http://poi.apache.org/spreadsheet/index.html)
	and is adapted to be used with ZK Spreadsheet component

---
zsstest
	A unit test set for zss.
	
---
zssdemo
	A simple Web example application that demonstrate the power of ZK Spreadsheet.
	
---
zssdemos
	The Java EE version of the zssdemo