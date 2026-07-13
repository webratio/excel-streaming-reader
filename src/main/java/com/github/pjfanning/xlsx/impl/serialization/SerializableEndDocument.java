package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.EndDocument;

public class SerializableEndDocument extends SerializableXMLEvent implements EndDocument {

    private static final long serialVersionUID = 1L;

    public SerializableEndDocument(EndDocument document) {
        super(document);
    }

    @Override
    public boolean isEndDocument() {
        return true;
    }

    @Override
    public String toString() {
        return "";
    }
}
