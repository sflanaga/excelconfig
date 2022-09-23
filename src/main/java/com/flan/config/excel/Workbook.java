package com.flan.config.excel;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.zip.ZipFile;

public class Workbook extends DefaultHandler {

    public HashMap<String, String> sheetNames = new HashMap<>();
    public final StringBuilder bld = new StringBuilder(256);
    public final HashMap<String, ArrayList<ArrayList<String>>> tableData = new HashMap<>();
    public final HashMap<String, TableData> tableDataByName = new HashMap<>();

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        bld.setLength(0);
        if (qName.equals("sheet")) {
            //  <sheet name="ks2group" sheetId="6" r:id="rId2"/>
            String n = attributes.getValue("name");
            String ridx = attributes.getValue("r:id").substring(3);
            sheetNames.put(n, ridx);
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        bld.append(ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
    }


    public Workbook(Path xlsxPath) throws IOException, ParserConfigurationException, SAXException {
        ZipFile zipFile = new ZipFile(xlsxPath.toFile(), ZipFile.OPEN_READ);
        long start = System.currentTimeMillis();
        try {

            // read them in an order that works for us
            try {
                SharedString ss = new SharedString();

                SAXParserFactory saxParserFactory = SAXParserFactory.newInstance();
                SAXParser saxParser = saxParserFactory.newSAXParser();
                saxParser.parse(zipFile.getInputStream(zipFile.getEntry("xl/workbook.xml")), this);
                saxParser.parse(zipFile.getInputStream(zipFile.getEntry("xl/sharedStrings.xml")), ss);

                ArrayList<Sheet> sheets = new ArrayList<>();
                // for each sheet we read - look for a .rels file to find associated tables within
                // we can bounce around with in the zip reading what we need as we go
                for(var e: this.sheetNames.entrySet()) {
                    var aSheet = new Sheet(e.getKey(), ss);
                    var sheetname = "xl/worksheets/sheet" + e.getValue() + ".xml";
                    var sheetEntry = zipFile.getEntry(sheetname);
                    if ( sheetEntry == null )
                        throw new RuntimeException("Excel parse error: cannot find expect sheet file: " + sheetname);
                    saxParser.parse(zipFile.getInputStream(sheetEntry), aSheet);
                    sheets.add(aSheet);
                    // find relationship file for tables references
                    String relfilename = "xl/worksheets/_rels/sheet" + e.getValue() + ".xml.rels";
                    var rele = zipFile.getEntry(relfilename);
                    if ( rele != null ) {
                        SheetRels rels = new SheetRels();
                        saxParser.parse(zipFile.getInputStream(rele), rels);
                        // add table references to this sheet
                        for(String fn: rels.tableFileName) {
                            var tableE = zipFile.getEntry("xl/tables/" + fn);
                            Table table = new Table();
                            saxParser.parse(zipFile.getInputStream(tableE), table);
                            aSheet.data().tablesRanges.putAll(table.tables);
                        }
                        extractTableData(aSheet);
                    }
                }
            } catch (ParserConfigurationException | SAXException e) {
                throw new RuntimeException("Sax error during parsing contents of " + xlsxPath.toString() + " reason: " + e.getLocalizedMessage(), e);
            }

        } finally {
            if (zipFile != null)
                zipFile.close();
        }
    }

    public void extractTableData(Sheet sheet) {
        for (var e : sheet.data().tablesRanges.entrySet()) {
            String tableName = e.getKey();
            var range = e.getValue();
            var tdata = new TableData(sheet.displaySheetName, tableName, range);
            tableDataByName.put(tableName, tdata);
            tdata.mapSheet(sheet);
        }
    }

    public static void main(String[] args) {
        try {
            long start = System.currentTimeMillis();
            Workbook wb = new Workbook(Path.of(args[0]));
            long end = System.currentTimeMillis();
            System.out.println("finished in " + (end-start) +"ms");

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
