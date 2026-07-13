package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.Characters;

public class SerializableCharacters extends SerializableXMLEvent implements Characters {

    private static final long serialVersionUID = 1L;

    private final String data;
    private final boolean cdata;
    private final boolean ignorableWhiteSpace;
    private final boolean whiteSpace;

    public SerializableCharacters(Characters characters) {
        super(characters);
        
        this.data = characters.getData();
        this.cdata = characters.isCData();
        this.ignorableWhiteSpace = characters.isIgnorableWhiteSpace();
        this.whiteSpace = characters.isWhiteSpace();
    }

    @Override
    public String getData() {
        return data;
    }

    @Override
    public boolean isCData() {
        return cdata;
    }

    @Override
    public boolean isIgnorableWhiteSpace() {
        return ignorableWhiteSpace;
    }

    @Override
    public boolean isWhiteSpace() {
        return whiteSpace;
    }

    @Override
    public Characters asCharacters() {
        return this;
    }

    @Override
    public boolean isCharacters() {
        return true;
    }

    @Override
    public String toString() {
        if (cdata) {
            return "<![CDATA[" + data + "]]>";
        }
        return data;
    }
}
