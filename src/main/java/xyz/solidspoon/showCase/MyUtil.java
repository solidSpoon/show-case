package xyz.solidspoon.showCase;

import cn.hutool.core.util.StrUtil;

public class MyUtil {
    public static String nullBlank(Object obj) {
        return obj == null ? "" : obj.toString();
    }
    public static String blankSpace(Object obj, int spaceLength) {
        return (obj == null || StrUtil.isBlank(obj.toString())) ? String.format("%" + spaceLength + "s", "") : obj.toString();
    }
}
