package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.NotationDeclaration;

public class SerializableNotationDeclaration extends SerializableXMLEvent implements NotationDeclaration {

    private static final long serialVersionUID = 1L;

    private final String name;
    private final String publicId;
    private final String systemId;

    public SerializableNotationDeclaration(NotationDeclaration notation) {
        super(notation);

        this.name = notation.getName();
        this.publicId = notation.getPublicId();
        this.systemId = notation.getSystemId();
    }

    @Override
    public String getName() {
        return name;
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
        StringBuilder sb = new StringBuilder("<!NOTATION ")
                .append(name);

        if (publicId != null) {
            sb.append(" PUBLIC \"")
              .append(publicId)
              .append("\"");
        } else if (systemId != null) {
            sb.append(" SYSTEM \"")
              .append(systemId)
              .append("\"");
        }

        return sb.append(">")
                 .toString();
    }
}