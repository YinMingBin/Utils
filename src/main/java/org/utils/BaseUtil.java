package org.utils;

import org.apache.commons.collections4.CollectionUtils;
import org.custom.collection.MatchingCollectList;
import org.custom.collection.MatchingList;
import org.custom.function.TypeFunction;
import org.custom.function.VoidFunction;
import org.utils.roughly.Matching;
import org.utils.roughly.MatchingCollect;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 操作对象工具类
 *
 * @author Administrator
 */
public class BaseUtil {
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
     * 对象转map
     *
     * @param item 要转换的对象
     * @return 转换后的map
     */
    public static Map<String, Object> objectToMap(Object item) {
        Map<String, Object> map = new HashMap<>(10);
        try {
            Class<?> aClass = item.getClass();
            for (Field field : aClass.getDeclaredFields()) {
                field.setAccessible(true);
                String fieldName = field.getName();
                Object value = field.get(item);
                map.put(fieldName, value);
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return map;
    }

    /**
     * 对象指定属性转map
     *
     * @param item      要转换的对象
     * @param functions 要转换的属性
     * @return 转换后的map
     */
    @SafeVarargs
    public static <T, R> Map<String, Object> objectToMap(T item, TypeFunction<T, R>... functions) {
        Map<String, Object> map = new HashMap<>(10);
        for (TypeFunction<T, R> function : functions) {
            map.put(TypeFunction.getLambdaColumnName(function), function.apply(item));
        }
        return map;
    }

    /**
     * 将obj1的值赋给obj2（通过属性名调用get/set方法）
     *
     * @param obj1 取值对象
     * @param obj2 赋值对象
     */
    public static void objectSetValue(Object obj1, Object obj2) {
        Class<?> aClass = obj1.getClass();
        Class<?> aClass1 = obj2.getClass();
        Field[] fields = aClass.getDeclaredFields();
        for (Field field : fields) {
            String name = field.getName();
            char[] chars = name.toCharArray();
            chars[0] -= 32;
            String name2 = String.valueOf(chars);
            try {
                Method declaredMethod1 = aClass.getDeclaredMethod("get" + name2);
                Object invoke = declaredMethod1.invoke(obj1);
                Method declaredMethod = aClass1.getDeclaredMethod("set" + name2, declaredMethod1.getReturnType());
                declaredMethod.invoke(obj2, invoke);
            } catch (IllegalAccessException | InvocationTargetException e) {
                System.out.println("属性" + name + "的方法调用失败!!!");
                e.printStackTrace();
            } catch (NoSuchMethodException e) {
                System.out.println("属性" + name + "没有get/set方法!!!");
                e.printStackTrace();
            }
        }
    }

    /**
     * 将obj1的某些值赋给obj2
     * 通过Function获取obj1的值，再获取Function对应的set方法名，通过方法名调用obj2的方法赋值
     *
     * @param obj1      取值对象
     * @param obj2      赋值对象
     * @param functions 要赋值的属性的get方法
     */
    @SafeVarargs
    public static <T, R> void objectSetValue(T obj1, Object obj2, TypeFunction<T, R>... functions) {
        Class<?> aClass = obj2.getClass();
        for (TypeFunction<T, R> function : functions) {
            R apply = function.apply(obj1);
            String setName = TypeFunction.getToSet(function);
            try {
                Method declaredMethod = aClass.getDeclaredMethod(setName, apply.getClass());
                declaredMethod.invoke(obj2, apply);
            } catch (NoSuchMethodException e) {
                System.out.println("没有" + setName + "方法");
                e.printStackTrace();
            } catch (IllegalAccessException | InvocationTargetException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 对象集一一对应设置值
     *
     * @param obj1      赋值对象集
     * @param obj2      取值对象集
     * @param functions 要赋值的属性
     * @param <T>       赋值对象类型
     * @param <R>       值类型
     */
    @SafeVarargs
    public static <T, R> void objectListSetValue(List<T> obj1, List<?> obj2, TypeFunction<T, R>... functions) {
        if (CollectionUtils.isNotEmpty(obj1) && CollectionUtils.isNotEmpty(obj2)) {
            Class<?> aClass = obj2.get(0).getClass();
            T t = obj1.get(0);
            for (TypeFunction<T, R> function : functions) {
                String setName = TypeFunction.getToSet(function);
                try {
                    Method declaredMethod = aClass.getDeclaredMethod(setName, function.apply(t).getClass());
                    for (int i = 0; i < obj2.size(); i++) {
                        R apply = function.apply(obj1.get(i));
                        declaredMethod.invoke(obj2.get(i), apply);
                    }
                } catch (NoSuchMethodException e) {
                    System.out.println("没有" + setName + "方法");
                    e.printStackTrace();
                } catch (IllegalAccessException | InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 匹配赋值
     *
     * @param obj1      赋值对象集
     * @param obj2      取值对象集
     * @param fun1      赋值对象匹配字段
     * @param fun2      取值对象匹配字段
     * @param functions 要赋值的属性
     * @param <T>       赋值对象类型
     * @param <I>       取值对象类型
     * @param <R>       匹配字段类型
     */
    public static <T, I, R> void matching(List<T> obj1, List<I> obj2, Function<T, R> fun1, Function<I, R> fun2,
                                          List<Matching<T, I>> functions) {
        matching(obj1, obj2, fun1, fun2, (t, i) -> functions.forEach(fun -> fun.setValue(t, i)));
    }

    public static <T, I, R> void matching(List<T> obj1, List<I> obj2, Function<T, R> fun1, Function<I, R> fun2,
                                          MatchingList<T, I> matchingList) {
        matching(obj1, obj2, fun1, fun2, (t, i) -> matchingList.forEach(fun -> fun.setValue(t, i)));
    }

    public static <T, I, R> void matching(List<T> obj1, List<I> obj2, Function<T, R> fun1, Function<I, R> fun2, VoidFunction.Two<T, I> function) {
        if (obj1 != null && obj1.size() > 0 && obj2 != null && obj2.size() > 0) {
            Map<R, I> riMap = toMapKey1(obj2, fun2);
            for (T t : obj1) {
                if (t != null) {
                    I i = riMap.get(fun1.apply(t));
                    if (i != null) {
                        function.apply(t, i);
                    }
                }
            }
        }
    }

    /**
     * 匹配赋值
     *
     * @param obj1            赋值对象集
     * @param matchingCollect 匹配取值对象
     * @param <T>             赋值对象类型
     */
    public static <T> void matching(List<T> obj1, MatchingCollect<T> matchingCollect) {
        if (obj1 != null && obj1.size() > 0 && matchingCollect != null) {
            for (T t : obj1) {
                if (t != null) {
                    matchingCollect.setValue(t);
                }
            }
        }
    }

    /**
     * 多个匹配赋值
     *
     * @param obj1             赋值对象集
     * @param matchingCollects 多个匹配取值对象
     * @param <T>              赋值对象类型
     */
    @SafeVarargs
    public static <T> void matching(List<T> obj1, MatchingCollect<T>... matchingCollects) {
        if (obj1 != null && obj1.size() > 0 && matchingCollects != null && matchingCollects.length > 0) {
            for (T t : obj1) {
                if (t != null) {
                    for (MatchingCollect<T> matchingCollect : matchingCollects) {
                        matchingCollect.setValue(t);
                    }
                }
            }
        }
    }

    /**
     * 多个匹配赋值
     *
     * @param obj1                赋值对象集
     * @param matchingCollectList 匹配取值对象集
     * @param <T>                 赋值对象类型
     */
    public static <T> void matching(List<T> obj1, MatchingCollectList<T> matchingCollectList) {
        if (obj1 != null && obj1.size() > 0 && matchingCollectList != null && matchingCollectList.isNotEmpty()) {
            for (T t : obj1) {
                if (t != null) {
                    matchingCollectList.forEach(m -> m.setValue(t));
                }
            }
        }
    }

}
