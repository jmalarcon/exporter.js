# exporter.js
A simple approach to allow downloading of HTML excerpts as several file types: .xls, .csv, .doc or .html

Including this file in your page will create a global `Exporter` object with a single method `export`.

This method's signature is:

```
Exporter.export(DOMElmt, fileName, dataName);
```

it takes a DOM element whose contents you want to export, the file name for the resulting file (it uses the file extension to determine the format) and optionally a name to be used in Excel files as the Workbook name of the exported .xls.

It supports: **Excel** (.xls), **CSV** (.csv, with `,` as a separator), **Word** (.doc) and **HTML** (.html) as file formats.

It works on any modern browser except Microsoft Edge, and it supports older versions of Internet Explorer (I've only tested it down to version 8, but it probably works in earlier versions too).

To use it just add this code to the `click` event of any link or button:

```
Exporter.export(data, 'courses.xls', 'Courses');return false;
```

It supports non-english characters in the data. The included sample .html file is in Spanish so it has plenty of those characters for you to test.

HTH!