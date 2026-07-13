package com.github.pjfanning.xlsx.impl.serialization;

import javax.xml.stream.events.Comment;

public class SerializableComment extends SerializableXMLEvent implements Comment {

    private static final long serialVersionUID = 1L;

    private final String text;

    public SerializableComment(Comment comment) {
        super(comment);
        
        this.text = comment.getText();
    }

    @Override
    public String getText() {
        return text;
    }

    @Override
    public String toString() {
        return "<!--" + text + "-->";
    }
}
