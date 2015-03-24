package com.pojo;

public class TestCase {

    private String testCaseName;
    private int no_of_pass;
    private int no_of_fail;

    public TestCase(String vName, int vPass, int vFail) {
        this.testCaseName = vName;
        this.no_of_pass = vPass;
        this.no_of_fail = vFail;
    }

    public int getNo_of_fail() {
        return no_of_fail;
    }

    public int getNo_of_pass() {
        return no_of_pass;
    }

    public String getTestCaseName() {
        return testCaseName;
    }
}
