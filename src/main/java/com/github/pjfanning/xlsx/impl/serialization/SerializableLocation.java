package com.github.pjfanning.xlsx.impl.serialization;

import java.io.Serializable;

import javax.xml.stream.Location;

public class SerializableLocation implements Location, Serializable {

    private static final long serialVersionUID = 1L;

    private final int lineNumber;
    private final int columnNumber;
    private final int characterOffset;
    private final String publicId;
    private final String systemId;

    public SerializableLocation(Location location) {
        this.lineNumber = location.getLineNumber();
        this.columnNumber = location.getColumnNumber();
        this.characterOffset = location.getCharacterOffset();
        this.publicId = location.getPublicId();
        this.systemId = location.getSystemId();
    }

    @Override
    public int getLineNumber() {
        return lineNumber;
    }

    @Override
    public int getColumnNumber() {
        return columnNumber;
    }

    @Override
    public int getCharacterOffset() {
        return characterOffset;
    }

    @Override
    public String getPublicId() {
        return publicId;
    }

    @Override
    public String getSystemId() {
        return systemId;
    }

    @Override
    public String toString() {
        return "SerializableLocation{" +
                "lineNumber=" + lineNumber +
                ", columnNumber=" + columnNumber +
                ", characterOffset=" + characterOffset +
                ", publicId='" + publicId + '\'' +
                ", systemId='" + systemId + '\'' +
                '}';
    }
}
