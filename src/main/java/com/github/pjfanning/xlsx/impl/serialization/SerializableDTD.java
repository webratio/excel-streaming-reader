package com.github.pjfanning.xlsx.impl.serialization;

import java.util.List;

import javax.xml.stream.events.DTD;
import javax.xml.stream.events.EntityDeclaration;
import javax.xml.stream.events.NotationDeclaration;

import org.apache.commons.collections4.list.UnmodifiableList;

public class SerializableDTD extends SerializableXMLEvent implements DTD {

    private static final long serialVersionUID = 1L;

    private final String documentTypeDeclaration;
    private final UnmodifiableList<EntityDeclaration> entities;
    private final UnmodifiableList<NotationDeclaration> notations;

    public SerializableDTD(DTD dtd) {
        super(dtd);

        this.documentTypeDeclaration = dtd.getDocumentTypeDeclaration();
        this.entities = new UnmodifiableList<>((List<EntityDeclaration>)dtd.getEntities());
        this.notations = new UnmodifiableList<>((List<NotationDeclaration>)dtd.getNotations());
    }

    @Override
    public String getDocumentTypeDeclaration() {
        return documentTypeDeclaration;
    }

    @Override
    public List<EntityDeclaration> getEntities() {
        return entities;
    }

    @Override
    public List<NotationDeclaration> getNotations() {
        return notations;
    }

    @Override
    public String toString() {
        return documentTypeDeclaration;
    }

    @Override
    public Object getProcessedDTD() {
        throw new UnsupportedOperationException("Unimplemented method 'getProcessedDTD'");
    }
}