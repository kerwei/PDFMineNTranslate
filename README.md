# PDFMineNTranslate

A Python script to read PDF files for Spanish keywords and retrieves the English translation for the matched string of text via the Microsoft Translator API

Requirements:
1. The "msmt" module written by Denis Papathanasiou, forked from https://gist.github.com/dpapathanasiou/2790853. It allows you to skip laying the groundwork to communicate with the Microsoft Translator API. "msmt.py" should be saved to the project root folder.
2. Keep all the PDF files in a subfolder at the root of the project folder. Update the name of the folder in pdfreadtranslate.py.
3. Create another subfolder with a desired name at the root of the project folder. This folder will hold the text version of the pdf files that contain the matched keywords.
4. At the end of the data mining process, an additional spreadsheet file will be created by the Python script to display the results. The result page would contain the following columns:
  i. "Text file"    Hyperlinks to the text files that were generated and kept in the folder described in Point 3.
  ii. "ES Match"    The excerpts of the Spanish content from the pdf files, in which the keywords are found.
  iii. "EN match"   The excerpts described above, translated to English.

Note:
If the pdf file is actually an image (i.e. a scanned/photographed copy of the original document), this script is unlikely to produce meaningful results.