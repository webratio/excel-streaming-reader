package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.Comment;
import javax.xml.stream.events.DTD;
import javax.xml.stream.events.EndDocument;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.EntityDeclaration;
import javax.xml.stream.events.EntityReference;
import javax.xml.stream.events.Namespace;
import javax.xml.stream.events.NotationDeclaration;
import javax.xml.stream.events.ProcessingInstruction;
import javax.xml.stream.events.StartDocument;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

public final class XMLEventUtils {

    private XMLEventUtils() {
        // utility class
    }

    public static SerializableXMLEvent toSerializable(XMLEvent event) {
        if (event == null) {
            return null;
        }
        if (event instanceof StartElement) {
            return new SerializableStartElement((StartElement) event);
        }
        if (event instanceof EndElement) {
            return new SerializableEndElement((EndElement) event);
        }
        if (event instanceof Characters) {
            return new SerializableCharacters((Characters) event);
        }
        if (event instanceof Attribute) {
            return new SerializableAttribute((Attribute) event);
        }
        if (event instanceof Namespace) {
            return new SerializableNamespace((Namespace) event);
        }
        if (event instanceof Comment) {
            return new SerializableComment((Comment) event);
        }
        if (event instanceof DTD) {
            return new SerializableDTD((DTD) event);
        }
        if (event instanceof EntityDeclaration) {
            return new SerializableEntityDeclaration((EntityDeclaration) event);
        }
        if (event instanceof EntityReference) {
            return new SerializableEntityReference((EntityReference) event);
        }
        if (event instanceof NotationDeclaration) {
            return new SerializableNotationDeclaration((NotationDeclaration) event);
        }
        if (event instanceof ProcessingInstruction) {
            return new SerializableProcessingInstruction((ProcessingInstruction) event);
        }
        if (event instanceof StartDocument) {
            return new SerializableStartDocument((StartDocument) event);
        }
        if (event instanceof EndDocument) {
            return new SerializableEndDocument((EndDocument) event);
        }
        throw new IllegalArgumentException(
                "Unsupported XMLEvent type: " + event.getClass().getName()
        );
    }
}
