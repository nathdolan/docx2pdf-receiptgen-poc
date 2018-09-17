
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;


public class Docx2PdfReceiptGenPoc {


    public static void main(String[] args) throws Exception {
        if (args.length != 2) {
            System.err.println("Provide file path to template docx and file path to output PDF as 2 arguments.");
            System.exit(-1);
        }
        new Docx2PdfReceiptGenPoc(args[0], args[1]);
    }


    public Docx2PdfReceiptGenPoc(String templateDocx, String pdfOut) throws Exception {
        while (true) {
            createPdf(templateDocx, pdfOut);
            Thread.sleep(3000);
        }
    }


    public void createPdf(String templateDocx, String pdfOut) throws Exception {

        long startTime = new Date().getTime();

        // Document loading (required)
        System.out.println("Loading file from " + templateDocx);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(templateDocx));
        System.out.println("Stage 1: " + (new Date().getTime() - startTime) + "ms");
        debugPrintDocxStructure(wordMLPackage.getMainDocumentPart());
        debugPrintDocxXml(wordMLPackage.getMainDocumentPart());

        // Set dynamic content
        writeSubstituteVars(wordMLPackage);
        writeDataTable(wordMLPackage.getMainDocumentPart());

        // Refresh the values of DOCPROPERTY fields
        FieldUpdater updater = new FieldUpdater(wordMLPackage);
        updater.update(true);

        OutputStream os = new java.io.FileOutputStream(pdfOut);
        FOSettings foSettings = Docx4J.createFOSettings();
        foSettings.setWmlPackage(wordMLPackage);
        System.out.println("Stage 2: " + (new Date().getTime() - startTime) + "ms");

        // Don't care what type of exporter you use
        Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
        System.out.println("Saved: " + pdfOut);

        // Clean up, so any ObfuscatedFontPart temp files can be deleted
        if (wordMLPackage.getMainDocumentPart().getFontTablePart() != null) {
            wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
        }

        System.out.println("End: " + (new Date().getTime() - startTime) + "ms");
    }


    private void writeSubstituteVars(WordprocessingMLPackage wordMLPackage) throws Exception {

        java.util.HashMap mappings = new java.util.HashMap();
        VariablePrepare.prepare(wordMLPackage);
        mappings.put("receiptId", "b09d69a5-e967-4f12-893b-e211df290a8d");
        mappings.put("businessEventEntity", "VAT return");
        mappings.put("date", "20 April 2018");
        mappings.put("time", "10:48");
        mappings.put("declarantName", "Nathan Dolan");
        mappings.put("declarantRole", "Director");
        mappings.put("declaration", "I declare that the information given in this form and accompanying documents is true and complete");
        wordMLPackage.getMainDocumentPart().variableReplace(mappings);
    }


    private void writeDataTable(MainDocumentPart documentPart) throws Exception {

        Text secHeaderText = (Text)XmlUtils.unwrap(documentPart.getJAXBNodesViaXPath("/w:document/w:body/w:p/w:r/w:t[contains(text(),'S1')]", false).get(0));
        P secHeaderP = (P)((R)secHeaderText.getParent()).getParent();
        Tbl tbl = (Tbl)XmlUtils.unwrap(documentPart.getJAXBNodesViaXPath("/w:document/w:body/w:tbl", false).get(0));
        Tr tr = (Tr)tbl.getContent().get(0);
        Tbl tblTemplate = (Tbl)XmlUtils.deepCopy(tbl);
        Tr trTemplate = (Tr)XmlUtils.deepCopy(tr);
        final int numSections = 3;
        final int numRows = 4;

        for (int section = 1; section <= numSections; section++) {

            if (section == 3)
                secHeaderText.setValue("Really really really really really long section " + section + " name");
            else
                secHeaderText.setValue("Section " + section + " name");

            for (int row = 1; row <= numRows; row++) {

                if (row == 2) // multi-line value example
                    setTableRowContent(tr, "Section " + section + " receipt data field " + row, "Address line 1", "Address line 2", "Town", "AB12 3CD");
                if (row == 3) // long value example
                    setTableRowContent(tr, "Section " + section + " receipt data field " + row, "Long long long long long long long long long long long long long field value");
                else // single value example
                    setTableRowContent(tr, "Section " + section + " receipt data field " + row, "Receipt data value " + row);

                if (row < numRows) {
                    // clone and add new table row
                    tr = XmlUtils.deepCopy(trTemplate);
                    tr.setParent(tbl);
                    tbl.getContent().add(tr);
                }
            }

            if (section < numSections) {

                int insertIdx = getTblIndex(tbl) + 1;
                Body body = (Body) tbl.getParent();

                // clone and insert new table under last table
                tbl = XmlUtils.deepCopy(tblTemplate);
                tbl.setParent(body);
                body.getContent().add(insertIdx, tbl);
                tr = (Tr)tbl.getContent().get(0);

                // clone and insert new section header P under last table
                secHeaderP = XmlUtils.deepCopy(secHeaderP);
                secHeaderText = (Text) XmlUtils.unwrap(((R) secHeaderP.getContent().get(0)).getContent().get(0));
                secHeaderP.setParent(body);
                body.getContent().add(insertIdx, secHeaderP);
            }
        }
    }


    private int getTblIndex(Tbl tbl) throws Exception {

        Body body = (Body)tbl.getParent();
        List<Object> content = body.getContent();
        for (int i = 0; i < content.size(); i++) {
            Object child = content.get(i);
            if (child == tbl || (child instanceof JAXBElement && XmlUtils.unwrap(child) == tbl))
                return i;
        }
        throw new Exception("Invalid state - cannot find given tbl in parent's content list");
    }


    private void setTableRowContent(Tr row, String fieldName, String... fieldValue) {

        setTableCell((Tc)XmlUtils.unwrap(row.getContent().get(0)), fieldName);
        setTableCell((Tc)XmlUtils.unwrap(row.getContent().get(1)), fieldValue);
    }


    private void setTableCell(Tc tc, String... textLines) {

        P valueP = (P) tc.getContent().get(0);

        for (int i = 0; i < textLines.length; i++) {

            R valueR = (R) valueP.getContent().get(0);
            Text valueText = (Text) XmlUtils.unwrap(valueR.getContent().get(0));
            valueText.setValue(textLines[i]);

            if (i < textLines.length - 1) {

                valueP = XmlUtils.deepCopy(valueP);
                valueP.setParent(tc);
                tc.getContent().add(valueP);
            }
        }
    }






    private void debugPrintDocxStructure(MainDocumentPart documentPart) {

        System.out.println();
        System.out.println("=======================================================================");
        System.out.println("DOCX JAXB STRUCTURE");
        System.out.println("=======================================================================");
        System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));
        System.out.println();
        System.out.println();

        org.docx4j.wml.Document wmlDocumentEl = documentPart.getJaxbElement();
        Body body = wmlDocumentEl.getBody();

        new TraversalUtil(body,

            new TraversalUtil.Callback() {

                String indent = "";

                public List<Object> apply(Object o) {

                    String wrapped = "";
                    if (o instanceof JAXBElement) wrapped =  " (wrapped in JAXBElement)";
                    o = XmlUtils.unwrap(o);
                    String text = "";
                    if (o instanceof org.docx4j.wml.Text)
                        text = ((org.docx4j.wml.Text) o).getValue();

                    System.out.println(indent + o.getClass().getName() + wrapped + "  \""
                            + text + "\"");
                    return null;
                }

                public boolean shouldTraverse(Object o) {
                    return true;
                }

                // Depth first
                public void walkJAXBElements(Object parent) {

                    indent += "    ";
                    List children = getChildren(parent);
                    if (children != null) {
                        for (Object o : children) {
                            this.apply(o);
                            // if its wrapped in javax.xml.bind.JAXBElement, get its
                            // value
                            o = XmlUtils.unwrap(o);
                            if (this.shouldTraverse(o)) {
                                walkJAXBElements(o);
                            }

                        }
                    }
                    indent = indent.substring(0, indent.length() - 4);
                }

                public List<Object> getChildren(Object o) {
                    return TraversalUtil.getChildrenImpl(o);
                }

            }
        );
    }


    private void debugPrintDocxXml(MainDocumentPart documentPart) {

        System.out.println();
        System.out.println("=======================================================================");
        System.out.println("DOCX XML");
        System.out.println("=======================================================================");
        System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));
        System.out.println();
        System.out.println();
    }

}


