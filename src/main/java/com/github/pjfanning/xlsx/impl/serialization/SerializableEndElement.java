package com.github.pjfanning.xlsx.impl.serialization;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.namespace.QName;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.Namespace;

import org.apache.commons.collections4.list.UnmodifiableList;

public class SerializableEndElement extends SerializableXMLEvent implements EndElement {

    private static final long serialVersionUID = 1L;

    private final QName name;
    private final UnmodifiableList<Namespace> namespaces;

    public SerializableEndElement(EndElement element) {
        super(element);

        this.name = element.getName();

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
    public Iterator<Namespace> getNamespaces() {
        return namespaces.iterator();
    }

    @Override
    public EndElement asEndElement() {
        return this;
    }

    @Override
    public boolean isEndElement() {
        return true;
    }

    @Override
    public String toString() {
        return "</" + name.getLocalPart() + ">";
    }
}
