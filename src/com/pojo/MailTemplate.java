package com.pojo;

import com.thoughtworks.xstream.annotations.XStreamAlias;

public class MailTemplate {

    private String to;
    private String cc;
    private String subject;
    private String body;
    private String[] to_arr;
    private String[] cc_arr;

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

    public String[] getCc_arr() {
        this.cc_arr = this.to.split(",");
        return cc_arr;
    }

    public String[] getTo_arr() {
        this.to_arr = this.cc.split(",");
        return to_arr;
    }
}
