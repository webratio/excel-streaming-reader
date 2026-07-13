package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.EntityDeclaration;

public class SerializableEntityDeclaration extends SerializableXMLEvent implements EntityDeclaration {

    private static final long serialVersionUID = 1L;

    private final String name;
    private final String replacementText;
    private final String notationName;
    private final String publicId;
    private final String systemId;
    private final String baseURI;

    public SerializableEntityDeclaration(EntityDeclaration entity) {
        super(entity);

        this.name = entity.getName();
        this.replacementText = entity.getReplacementText();
        this.notationName = entity.getNotationName();
        this.publicId = entity.getPublicId();
        this.systemId = entity.getSystemId();
        this.baseURI = entity.getBaseURI();
    }

    @Override
    public String getName() {
        return name;
    }

    @Override
    public String getReplacementText() {
        return replacementText;
    }

    @Override
    public String getNotationName() {
        return notationName;
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
    public String getBaseURI() {
        return baseURI;
    }

    @Override
    public String toString() {
        return "<!ENTITY " + name + ">";
    }
}
