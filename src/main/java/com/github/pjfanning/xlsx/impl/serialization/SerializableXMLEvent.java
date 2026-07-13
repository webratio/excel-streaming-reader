package com.github.pjfanning.xlsx.impl.serialization;

import java.io.IOException;
import java.io.Serializable;
import java.io.Writer;

import javax.xml.namespace.QName;
import javax.xml.stream.Location;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public abstract class SerializableXMLEvent implements XMLEvent, Serializable {

    private static final long serialVersionUID = 1L;

    private final int eventType;
    protected final SerializableLocation location;

    public SerializableXMLEvent(XMLEvent event) {
        this.eventType = event.getEventType();
        this.location = new SerializableLocation(event.getLocation());
    }

    @Override
    public int getEventType() {
        return eventType;
    }

    @Override
    public Location getLocation() {
        return location;
    }

    @Override
    public boolean isStartElement() {
        return eventType == XMLStreamConstants.START_ELEMENT;
    }

    @Override
    public boolean isAttribute() {
        return eventType == XMLStreamConstants.ATTRIBUTE;
    }

    @Override
    public boolean isNamespace() {
        return eventType == XMLStreamConstants.NAMESPACE;
    }

    @Override
    public boolean isEndElement() {
        return eventType == XMLStreamConstants.END_ELEMENT;
    }

    @Override
    public boolean isEntityReference() {
        return eventType == XMLStreamConstants.ENTITY_REFERENCE;
    }

    @Override
    public boolean isProcessingInstruction() {
        return eventType == XMLStreamConstants.PROCESSING_INSTRUCTION;
    }

    @Override
    public boolean isCharacters() {
        return eventType == XMLStreamConstants.CHARACTERS
                || eventType == XMLStreamConstants.CDATA
                || eventType == XMLStreamConstants.SPACE;
    }

    @Override
    public boolean isStartDocument() {
        return eventType == XMLStreamConstants.START_DOCUMENT;
    }

    @Override
    public boolean isEndDocument() {
        return eventType == XMLStreamConstants.END_DOCUMENT;
    }

    @Override
    public StartElement asStartElement() {
        if (!isStartElement()) {
            throw new IllegalStateException("Event is not a StartElement");
        }
        return (StartElement) this;
    }

    @Override
    public EndElement asEndElement() {
        if (!isEndElement()) {
            throw new IllegalStateException("Event is not an EndElement");
        }
        return (EndElement) this;
    }

    @Override
    public Characters asCharacters() {
        if (!isCharacters()) {
            throw new IllegalStateException("Event is not Characters");
        }
        return (Characters) this;
    }

    @Override
    public QName getSchemaType() {
        return null;
    }

    @Override
    public void writeAsEncodedUnicode(Writer writer) throws XMLStreamException {
        try {
            writer.write(toString());
        } catch (IOException e) {
            throw new XMLStreamException(e);
        }
    }

    @Override
    public String toString() {
        return "SerializableXMLEvent{" +
                "eventType=" + eventType +
                ", location=" + location +
                '}';
    }
}
