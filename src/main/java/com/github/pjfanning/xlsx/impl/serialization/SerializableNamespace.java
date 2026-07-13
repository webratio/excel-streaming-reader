package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.namespace.QName;
import javax.xml.stream.events.Namespace;

public class SerializableNamespace extends SerializableXMLEvent implements Namespace {

    private static final long serialVersionUID = 1L;

    private final String prefix;
    private final String namespaceURI;
    private final boolean defaultNamespace;

    public SerializableNamespace(Namespace namespace) {
        super(namespace);

        this.prefix = namespace.getPrefix();
        this.namespaceURI = namespace.getNamespaceURI();
        this.defaultNamespace = namespace.isDefaultNamespaceDeclaration();
    }

    @Override
    public String getPrefix() {
        return prefix;
    }

    @Override
    public String getNamespaceURI() {
        return namespaceURI;
    }

    @Override
    public boolean isDefaultNamespaceDeclaration() {
        return defaultNamespace;
    }

    @Override
    public QName getName() {
        if (defaultNamespace) {
            return new QName(
                    javax.xml.XMLConstants.XMLNS_ATTRIBUTE_NS_URI,
                    javax.xml.XMLConstants.XMLNS_ATTRIBUTE);
        }

        return new QName(
                javax.xml.XMLConstants.XMLNS_ATTRIBUTE_NS_URI,
                prefix,
                javax.xml.XMLConstants.XMLNS_ATTRIBUTE);
    }

    @Override
    public String getValue() {
        return namespaceURI;
    }

    @Override
    public String getDTDType() {
        return "CDATA";
    }

    @Override
    public boolean isSpecified() {
        return true;
    }

    @Override
    public boolean isNamespace() {
        return true;
    }

    @Override
    public String toString() {
        if (defaultNamespace) {
            return "xmlns=\"" + namespaceURI + "\"";
        }

        return "xmlns:" + prefix + "=\"" + namespaceURI + "\"";
    }
}
