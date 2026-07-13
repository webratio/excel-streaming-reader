package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.ProcessingInstruction;

public class SerializableProcessingInstruction extends SerializableXMLEvent implements ProcessingInstruction {

    private static final long serialVersionUID = 1L;

    private final String target;
    private final String data;

    public SerializableProcessingInstruction(ProcessingInstruction pi) {
        super(pi);

        this.target = pi.getTarget();
        this.data = pi.getData();
    }

    @Override
    public String getTarget() {
        return target;
    }

    @Override
    public String getData() {
        return data;
    }

    @Override
    public String toString() {
        if (data == null || data.isEmpty()) {
            return "<?" + target + "?>";
        }

        return "<?" + target + " " + data + "?>";
    }
}
