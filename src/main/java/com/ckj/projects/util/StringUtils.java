package com.ckj.projects.util;

/**
 * created by ChenKaiJu on 2018/11/2  13:55
 */
public class StringUtils {

    public static boolean isBlank(String src){
        if(src==null){
            return true;
        }
        if(src.trim().equals("")){
            return true;
        }
        return false;
    }
}
