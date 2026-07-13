package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.StartDocument;

public class SerializableStartDocument extends SerializableXMLEvent implements StartDocument {

    private static final long serialVersionUID = 1L;

    private final String systemId;
    private final String characterEncodingScheme;
    private final String version;
    private final boolean standalone;
    private final boolean standaloneSet;
    private final boolean encodingSet;

    public SerializableStartDocument(StartDocument document) {
        super(document);

        this.systemId = document.getSystemId();
        this.characterEncodingScheme = document.getCharacterEncodingScheme();
        this.version = document.getVersion();
        this.standalone = document.isStandalone();
        this.standaloneSet = document.standaloneSet();
        this.encodingSet = document.encodingSet();
    }

    @Override
    public String getSystemId() {
        return systemId;
    }

    @Override
    public String getCharacterEncodingScheme() {
        return characterEncodingScheme;
    }

    @Override
    public String getVersion() {
        return version;
    }

    @Override
    public boolean isStandalone() {
        return standalone;
    }

    @Override
    public boolean standaloneSet() {
        return standaloneSet;
    }

    @Override
    public boolean encodingSet() {
        return encodingSet;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder("<?xml");

        if (version != null) {
            sb.append(" version=\"")
              .append(version)
              .append("\"");
        }

        if (characterEncodingScheme != null) {
            sb.append(" encoding=\"")
              .append(characterEncodingScheme)
              .append("\"");
        }

        if (standaloneSet) {
            sb.append(" standalone=\"")
              .append(standalone ? "yes" : "no")
              .append("\"");
        }

        return sb.append("?>")
                 .toString();
    }
}
