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
}
