package com.github.pjfanning.xlsx.impl.serialization;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.namespace.NamespaceContext;
import javax.xml.namespace.QName;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Namespace;
import javax.xml.stream.events.StartElement;

import org.apache.commons.collections4.list.UnmodifiableList;

public class SerializableStartElement extends SerializableXMLEvent implements StartElement {

    private static final long serialVersionUID = 1L;

    private final QName name;
    private final UnmodifiableList<Attribute> attributes;
    private final UnmodifiableList<Namespace> namespaces;

    public SerializableStartElement(StartElement element) {
        super(element);
        this.name = element.getName();

        List<SerializableAttribute> attributes = new ArrayList<>();
        Iterator<?> attributesIterator = element.getAttributes();
        while (attributesIterator.hasNext()) {
            Attribute attribute = (Attribute) attributesIterator.next();
            attributes.add(new SerializableAttribute(attribute));
        }
        this.attributes = new UnmodifiableList<>(attributes);
        List<SerializableNamespace> namespaces = new ArrayList<>();
        Iterator<?> namespacesIterator = element.getNamespaces();
        while (namespacesIterator.hasNext()) {
            Namespace namespace = (Namespace) namespacesIterator.next();
            namespaces.add(new SerializableNamespace(namespace));
        }
        this.namespaces = new UnmodifiableList<>(namespaces);
    }

    @Override
    public QName getName() {
        return name;
    }

    @Override
    public Iterator<Attribute> getAttributes() {
        return attributes.iterator();
    }

    @Override
    public Iterator<Namespace> getNamespaces() {
        return namespaces.iterator();
    }

    @Override
    public Attribute getAttributeByName(QName qname) {
        for (Attribute attribute : attributes) {
            if (attribute.getName().equals(qname)) {
                return attribute;
            }
        }

        return null;
    }

    @Override
    public NamespaceContext getNamespaceContext() {
        return null;
    }

    @Override
    public String getNamespaceURI(String prefix) {
        for (Namespace namespace : namespaces) {
            if (namespace.getPrefix().equals(prefix)) {
                return namespace.getNamespaceURI();
            }
        }

        return null;
    }

    @Override
    public StartElement asStartElement() {
        return this;
    }

    @Override
    public boolean isStartElement() {
        return true;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();

        sb.append("<")
          .append(name.getLocalPart());

        for (Namespace namespace : namespaces) {
            sb.append(" ")
              .append(namespace);
        }

        for (Attribute attribute : attributes) {
            sb.append(" ")
              .append(attribute);
        }

        sb.append(">");

        return sb.toString();
    }
}
