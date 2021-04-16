package org.utils;

import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * list工具类
 *
 * @author YinMingBin
 */
public class ListUtil {

    /**
     * list转set
     *
     * @param list   数据集
     * @param mapper 获取值函数
     * @param <T>    对象类型
     * @param <R>    值类型
     * @return list
     */
    public static <T, R> List<R> toList(List<T> list, Function<? super T, ? extends R> mapper) {
        return list.stream().map(mapper).collect(Collectors.toList());
    }

    /**
     * list转set
     *
     * @param list   数据集
     * @param mapper 获取值函数
     * @param <T>    对象类型
     * @param <R>    值类型
     * @return set
     */
    public static <T, R> Set<R> toSet(List<T> list, Function<? super T, ? extends R> mapper) {
        return list.stream().map(mapper).collect(Collectors.toSet());
    }

    /**
     * list转map，不判断key重复
     *
     * @param objects   对象集
     * @param keyMapper 获取key函数
     * @param <T>       对象类型
     * @param <K>       key类型
     * @return map
     */
    public static <T, K> Map<K, T> toMap(Collection<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, Function.identity()));
        }
        return new HashMap<>(0);
    }

    /**
     * list转map，不判断key重复
     *
     * @param objects     对象集
     * @param keyMapper   获取key函数
     * @param valueMapper 获取value函数
     * @param <T>         对象类型
     * @param <K>         key类型
     * @param <V>         value类型
     * @return map
     */
    public static <T, K, V> Map<K, V> toMap(Collection<T> objects, Function<? super T, ? extends K> keyMapper,
                                            Function<? super T, ? extends V> valueMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, valueMapper));
        }
        return new HashMap<>(0);
    }

    /**
     * list转map，key重复保留前一位
     *
     * @param objects   对象集
     * @param keyMapper 获取key函数
     * @param <T>       对象类型
     * @param <K>       key类型
     * @return map
     */
    public static <T, K> Map<K, T> toMapKey1(Collection<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, Function.identity(), (key1, key2) -> key1));
        }
        return new HashMap<>(0);
    }

    /**
     * list转map，key重复保留前一位
     *
     * @param objects     对象集
     * @param keyMapper   获取key函数
     * @param valueMapper 获取value函数
     * @param <T>         对象类型
     * @param <K>         key类型
     * @param <V>         value类型
     * @return map
     */
    public static <T, K, V> Map<K, V> toMapKey1(Collection<T> objects, Function<? super T, ? extends K> keyMapper,
                                                Function<? super T, ? extends V> valueMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, valueMapper, (key1, key2) -> key1));
        }
        return new HashMap<>(0);
    }

    /**
     * list转map，key重复保留后一位
     *
     * @param objects   对象集
     * @param keyMapper 获取key函数
     * @param <T>       对象类型
     * @param <K>       key类型
     * @return map
     */
    public static <T, K> Map<K, T> toMapKey2(Collection<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, Function.identity(), (key1, key2) -> key2));
        }
        return new HashMap<>(0);
    }

    /**
     * list转map，key重复保留后一位
     *
     * @param objects     对象集
     * @param keyMapper   获取key函数
     * @param valueMapper 获取value函数
     * @param <T>         对象类型
     * @param <K>         key类型
     * @param <V>         value类型
     * @return map
     */
    public static <T, K, V> Map<K, V> toMapKey2(Collection<T> objects, Function<? super T, ? extends K> keyMapper,
                                                Function<? super T, ? extends V> valueMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, valueMapper, (key1, key2) -> key2));
        }
        return new HashMap<>(0);
    }

    /**
     * list分组
     *
     * @param objects   要分组的对象集
     * @param keyMapper 根据什么分组
     * @param <T>       分组对象类型
     * @param <K>       根据什么分组类型
     * @return 分组后数据
     */
    public static <T, K> Map<K, List<T>> groupingBy(List<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!objects.isEmpty()) {
            return objects.stream().collect(Collectors.groupingBy(keyMapper));
        }
        return new HashMap<>(0);
    }

    /**
     * list分组处理value
     *
     * @param objects     要分组的对象集
     * @param keyMapper   根据什么分组
     * @param valueMapper 分组后的value处理
     * @param <T>         分组对象类型
     * @param <K>         根据什么分组类型
     * @param <V>         value类型
     * @return 分组后数据
     */
    public static <T, K, V> Map<K, V> groupingBy(List<T> objects, Function<? super T, ? extends K> keyMapper,
                                                 Function<List<T>, ? extends V> valueMapper) {
        if (!objects.isEmpty()) {
            Map<? extends K, List<T>> collect = objects.stream().collect(Collectors.groupingBy(keyMapper));
            Map<K, V> map = new HashMap<>(collect.size());
            collect.forEach((key, value) -> {
                map.put(key, valueMapper.apply(value));
            });
            return map;
        }
        return new HashMap<>(0);
    }

    /**
     * 只保存相同数据（两边都存在的数据）
     *
     * @param list1 数据集1(两边都存在的数据会放到这里面)
     * @param list2 数据集2
     * @param <T>   数据类型
     */
    public static <T> void saveIdentical(Collection<T> list1, Collection<T> list2) {
        if (list1 != null) {
            Collection<T> list = new ArrayList<>();
            if (list2 != null) {
                for (T t : list1) {
                    if (list2.contains(t)) {
                        list.add(t);
                    }
                }
            }
            list1.clear();
            list1.addAll(list);
        }
    }
}
