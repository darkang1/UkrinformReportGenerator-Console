# Ukrinform Report Generator (Console Version)
A report-generating tool created by request of Ukrinform News Agency.

## Description
**Ukrinform Report Generator** (URG) is a tool designed to generate weekly report based on provided article files of format **.doc** or **.docx**.

User required to specify full path to the folder which contains article files he wants to parse and obtain generated report.

It can be done either by hardcoding full path inside the application (for cases if the same full path is always used) or to manually specify it inside application
by selecting appropriate option.

From the obtained article files, **URG** parses hyperlinks contained in the articles header and perform following information parsing from the article page on **Ukrinform** website:
- Article date
- Article header
- Article size (amount of characters without spaces)
- Article type
- Article exclusiveness

After obtaining all information required from each article page **URG** generates **Microsoft Word** report from hardcoded report template with automatically
generated header, dates, and filename based on the current system date.

## Requirements
- Each article should contain hyperlink to **Ukrinform** article webpage
  - Article webpage hyperlink should be the first hyperlink in article document
  - Only links within "ukrinform" domain are supported
  - If file contains no **Ukrinform** links then report will be generated specifying filename of the article and appropriate warning message
- Have system datetime installed to the appropriate date of week to which report is related to (for proper automatic generation of header, dates, and filename)
  - So it's a good idea to generate report at the end of each week (ex. **Sunday**) 

## Dependencies
- DocX
- HtmlAgilityPack
- System.IO.Packaging
