package com.flan.config.excel;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.HashMap;

public class Table extends DefaultHandler {

    public final StringBuilder bld = new StringBuilder(256);
    public final HashMap<String, CellRange> tables = new HashMap<>();

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        bld.setLength(0);
        switch(qName) {
            case "table":
                String n = attributes.getValue("displayName");
                String range = attributes.getValue("ref");
                CellRef[] refs=new CellRef[2];
                String[] refStrs = range.split(":");
                if ( refStrs.length != 2)
                    throw new RuntimeException("Cannot handle cell range format for: " + range + " for table display named: " + n);
                for (int i = 0; i < 2; i++) {
                    refs[i] = new CellRef(refStrs[i]);
                }
                tables.put(n, new CellRange(refs[0], refs[1]));
            break;
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        bld.append(ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
    }


}
