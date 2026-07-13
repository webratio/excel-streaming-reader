package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.namespace.QName;
import javax.xml.stream.events.Attribute;

public class SerializableAttribute extends SerializableXMLEvent implements Attribute {

    private static final long serialVersionUID = 1L;

    private final QName name;
    private final String value;
    private final String dtdType;
    private final boolean specified;

    public SerializableAttribute(Attribute attribute) {
        super(attribute);
        
        this.name = attribute.getName();
        this.value = attribute.getValue();
        this.dtdType = attribute.getDTDType();
        this.specified = attribute.isSpecified();
    }

    @Override
    public QName getName() {
        return name;
    }

    @Override
    public String getValue() {
        return value;
    }

    @Override
    public String getDTDType() {
        return dtdType;
    }

    @Override
    public boolean isSpecified() {
        return specified;
    }

    @Override
    public boolean isAttribute() {
        return true;
    }

    @Override
    public String toString() {
        return name + "=\"" + value + "\"";
    }
}
