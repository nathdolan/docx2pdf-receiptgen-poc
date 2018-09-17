# docx2pdf-receiptgen-poc

Proof of concept for manipulation of docx template and creation of PDF from docx using open-source (*not* iText).

## To run from built JAR
* Ensure Java 1.8+ is installed
* Copy docx2pdf-receiptgen-poc-1.0-SNAPSHOT-jar-with-dependencies.jar to a local folder
* From command line change to directory containing the above file and enter:
`java -cp ./docx2pdf-receiptgen-poc-1.0-SNAPSHOT-jar-with-dependencies.jar Docx2PdfReceiptGenPoc <path to docx template file> <path to the PDF file to create>`
