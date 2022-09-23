package com.flan.config.excel;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SheetRels  extends DefaultHandler {

    ArrayList<String> tableFileName = new ArrayList<>();
    static Pattern p = Pattern.compile(".*/(table\\d+.xml)");

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if (qName.equals("Relationship"))
            if ( attributes.getValue("Type").endsWith("relationships/table") ) {
                String tstr = attributes.getValue("Target");
                Matcher m = p.matcher(tstr);
                if ( m.matches() )
                    tableFileName.add(m.group(1));
            }
    }
}
