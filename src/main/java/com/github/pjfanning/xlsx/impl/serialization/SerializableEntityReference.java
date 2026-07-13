package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.EntityDeclaration;
import javax.xml.stream.events.EntityReference;

public class SerializableEntityReference extends SerializableXMLEvent implements EntityReference {

    private static final long serialVersionUID = 1L;

    private final String name;
    private final SerializableEntityDeclaration declaration;

    public SerializableEntityReference(EntityReference entityReference) {
        super(entityReference);

        this.name = entityReference.getName();

        EntityDeclaration declaration = entityReference.getDeclaration();
        this.declaration = declaration != null
                ? new SerializableEntityDeclaration(declaration)
                : null;
    }

    @Override
    public String getName() {
        return name;
    }

    @Override
    public EntityDeclaration getDeclaration() {
        return declaration;
    }

    @Override
    public String toString() {
        return "&" + name + ";";
    }
}
