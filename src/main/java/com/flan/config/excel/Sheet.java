package com.flan.config.excel;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public class Sheet extends DefaultHandler {
    private final StringBuilder bld = new StringBuilder(256);

    public final String displaySheetName;
    private final SheetData data = new SheetData();

    private final SharedString ss;

    public Sheet(String displaySheetName, SharedString ss) {
        this.displaySheetName = displaySheetName;
        this.ss = ss;
    }

    private int row = -1;
    private int col = -1;
    private char type = '0';
    private Row currRow = null;
    private String currRef = null;

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        bld.setLength(0);
        if (qName.equals("c")) {
            String r = attributes.getValue("r");
            currRef = r;
            populateRowCol(r);
            String t = attributes.getValue("t");
            if (t == null)
                type = '0';
            else {
                switch (t) {
                    case "s":
                        type = 'l'; // lookup in string table
                        break;
                    case "str":
                        type = 's'; // literal string in <v>
                        break;
                    case "b":
                        type = 'b'; // boolean TRUE/FALSE
                        break;
                    case "e":
                        type = 's';
                        break;
                    default:
                        throw new RuntimeException("Unsupported type: " + t + " at sheet: " + displaySheetName + " at cell: " + r);
                }
            }
        } else if (qName.equals("row")) {
            int r = Integer.valueOf(attributes.getValue("r"));
            currRow = new Row(r);
            data.rows.add(currRow);
        }
    }


    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        bld.append(ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        switch (qName) {
            case "v":
                String v = bld.toString().trim();
                switch (type) {
                    case 'l':
                        addValue(ss.sharedStrings.get(Integer.valueOf(v)));
                        break;
                    case '0':
                    case 's':
                        addValue(v);
                        break;
                    case 'b':
                        if (v.charAt(0) == '0')
                            addValue("false");
                        else
                            addValue("true");

                }
                break;
            case "c":
                col = -1;
                break;
            case "r":
                row = -1;
                break;
        }
    }

    private void addValue(String v) {
        CellData cd = new CellData(v, col);
        if (currRow.row != row)
            throw new RuntimeException("Row # mismatch between reference in internal count for the row, ref = " + currRef + " vs " + currRow.row);
        currRow.cols.add(cd);
    }

    private void populateRowCol(String s) {
        int i = 0;
        for (i = 0; i < s.length(); i++) {
            if (Character.isDigit(s.charAt(i)))
                break;
        }
        String digs = s.substring(i);
        col = lettersToRow(s.substring(0, i));
        row = Integer.valueOf(digs);
    }

    private static int lettersToRow(String s) {
        int v = 0;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            v *= 26;
            if (c >= 'A' && c <= 'Z')
                v += ((int) s.charAt(i)) - ((int) 'A') + 1;
            else
                throw new RuntimeException("Bad cell reference letters: " + s);
        }
        return v;
    }

    public SheetData data() { return data; }
}



