package org.custom.collection;

import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Function;

/**
 * 处理前端请求所有数据Map集
 *
 * @author Administrator
 */
public class QueryMap {
    private final Map<String, Function<Object, ?>> map = new HashMap<>(10);

    public QueryMap add(String key) {
        map.put(key, val -> val);
        return this;
    }

    public QueryMap add(String key, Function<Object, ?> value) {
        map.put(key, value);
        return this;
    }

    public QueryMap add(String key, Function<Object, ?>... values) {
        Function<Object, ?> value = val -> {
            for (Function<Object, ?> function : values) {
                val = function.apply(val);
            }
            return val;
        };
        map.put(key, value);
        return this;
    }

    public void forEach(BiConsumer<String, Function<Object, ?>> action) {
        map.forEach(action);
    }

    public Integer getSize() {
        return map.size();
    }
}