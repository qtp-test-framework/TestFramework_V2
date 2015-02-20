package com.pojo;

import com.thoughtworks.xstream.annotations.XStreamAlias;

public class MailTemplate {

    private String to;
    private String cc;
    private String subject;
    private String body;

    public MailTemplate(String vTo, String vCC, String vSub, String vBody) {
        this.to = vTo;
        this.cc = vCC;
        this.subject = vSub;
        this.body = vBody;
    }

    public String getBody() {
        return body;
    }

    public String getCc() {
        return cc;
    }

    public String getSubject() {
        return subject;
    }

    public String getTo() {
        return to;
    }
}
