package org.custom.collection;

import org.utils.object.MatchingKey;

import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 同数据多个匹配赋值list
 * @author Administrator
 */
public class MatchingKeyList<T, I>{
    private final Map<Function<I, ?>, List<? extends MatchingKey<T, I, ?>>> map = new HashMap<>();

    public <R> MatchingKeyList(Function<I, R> key, List<MatchingKey<T, I, R>> value) {
        map.put(key, value);
    }

    @SafeVarargs
    public <R> MatchingKeyList(Function<I, R> key, MatchingKey<T, I, R>... values) {
        map.put(key, Arrays.stream(values).collect(Collectors.toList()));
    }

    public <R> MatchingKeyList<T, I> add(Function<I, R> key, List<MatchingKey<T, I, R>> value){
        map.put(key, value);
        return this;
    }

    @SafeVarargs
    public final <R> MatchingKeyList<T, I> add(Function<I, R> key, MatchingKey<T, I, R>... values){
        map.put(key, Arrays.stream(values).collect(Collectors.toList()));
        return this;
    }

    public void forEach(BiConsumer<Function<I, ?>, List<? extends MatchingKey<T, I, ?>>> action) {
        map.forEach(action);
    }

    /**
     * 判断是否有值
     * @return 是否有值
     */
    public boolean isNotEmpty() {
        return !map.isEmpty();
    }

}
