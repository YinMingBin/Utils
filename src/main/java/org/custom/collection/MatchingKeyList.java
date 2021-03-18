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

    /**
     * 添加匹配赋值
     * @param key 取值匹配值函数
     * @param value 匹配赋值数据集
     * @param <R> 匹配值类型
     * @return 本身
     */
    public <R> MatchingKeyList<T, I> add(Function<I, R> key, List<MatchingKey<T, I, R>> value){
        map.put(key, value);
        return this;
    }

    /**
     * 添加匹配赋值
     * @param key 取值匹配值函数
     * @param values 多个匹配赋值数据
     * @param <R> 匹配值类型
     * @return 本身
     */
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
