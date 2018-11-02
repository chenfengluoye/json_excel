package com.ckj.projects.entity;

/**
 * created by ChenKaiJu on 2018/11/2  14:50
 */
public class Entry implements Comparable<Entry> {

    public String key;

    public  Object value;

    public Entry(String key,Object value){
        this.key=key;
        this.value=value;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    @Override
    public int compareTo(Entry o) {
        int uslenth=this.key.length();
        int otherlen=o.key.length();
        int result=uslenth-otherlen;
        if(result!=0){
            return result;
        }
        return this.key.compareTo(o.key);
    }
}
