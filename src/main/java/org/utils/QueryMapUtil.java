package org.utils;

import org.springframework.util.StringUtils;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Function;

/**
 * 处理请求的所有参数
 * @author Administrator
 */
public class QueryMapUtil {

    public static Map<String, Object> dispose(Map<String, Object> queryMap, QueryMap functionMap){
        Map<String, Object> map = new HashMap<>(functionMap.getSize());
        functionMap.forEach((key, value) -> {
            int i = key.indexOf("-");
            String name = key;
            if(i >= 0) {
                String keyCopy = key;
                key = keyCopy.substring(0, i);
                name = keyCopy.substring(i + 1);
            }
            Object o = queryMap.get(key);
            if(o != null) {
                Object apply = value.apply(o);
                if(apply != null){
                    map.put(name, apply);
                }
            }
        });
        return map;
    }

    /**
     * 时间戳转时间
     * @param dateTime 时间戳
     * @return 时间
     */
    public static Date dateTimeToDate(Object dateTime){
        if(dateTime != null) {
            String s = dateTime.toString();
            if(!StringUtils.isEmpty(s)){
                return new Date(Long.parseLong(s));
            }
        }
        return null;
    }

    /**
     * 转字符串
     * @param obj 对象
     * @return 字符串
     */
    public static String toString(Object obj) {
        if(obj != null) {
            return obj.toString();
        }
        return null;
    }

    /**
     * 转Long
     * @param obj 对象
     * @return Long
     */
    public static Long toLong(Object obj) {
        if(obj != null) {
            return Long.valueOf(obj.toString());
        }
        return null;
    }

    /**
     * 转Integer
     * @param obj 对象
     * @return Integer
     */
    public static Integer toInteger(Object obj) {
        if(obj != null) {
            return Integer.valueOf(obj.toString());
        }
        return null;
    }

    public static class QueryMap{
        private final Map<String, Function<Object, ?>> map = new HashMap<>(10);
        private Integer size = 0;

        public QueryMap add(String key){
            map.put(key, val -> val);
            size++;
            return this;
        }
        public QueryMap add(String key, Function<Object, ?> value){
            map.put(key, value);
            size++;
            return this;
        }
        public QueryMap add(String key, Function<Object, ?>... values){
            Function<Object, ?> value = val -> {
                for (Function<Object, ?> function : values) {
                    val = function.apply(val);
                }
                return val;
            };
            map.put(key, value);
            size++;
            return this;
        }

        public void forEach(BiConsumer<String, Function<Object, ?>> action) {
            map.forEach(action);
        }

        public Integer getSize() {
            return size;
        }
    }
}
