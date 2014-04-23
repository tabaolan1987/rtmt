package com.cmg.hipspot.util;

/**
 * Created by lantb on 2014-04-21.
 */
public class StringUtil {

    public static Object isNull(final Object o, final Object dflt) {
        if (o == null) {
            return dflt;
        } else {
            return o;
        }
    }
	
	public static String List2String(ArrayList<String> list){
       if(list.size() > 0 ){
           String temp = new String();
           for(String s : list){
               temp+= s +"|";
               System.out.println("temp : " + temp);
           }
           return temp;
       }else{
           return null;
       }
    }
}
