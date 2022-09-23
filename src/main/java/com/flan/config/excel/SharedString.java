package com.flan.config.excel;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;

public class SharedString extends DefaultHandler {

    final ArrayList<String> sharedStrings = new ArrayList<>(100);
    private final StringBuilder bld = new StringBuilder(256);
    private final StringBuilder currStr = new StringBuilder(256);

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        bld.setLength(0);
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        bld.append(ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        if ( qName.equals("t") )
            currStr.append(bld);
        else if ( qName.equals("si") ) {
            sharedStrings.add(currStr.toString().intern());
            currStr.setLength(0);
        }
    }
}
