package org.base;

import org.apache.commons.collections4.CollectionUtils;
import org.function.TypeFunction;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 操作对象工具类
 * @author Administrator
 */
public class BaseUtil {
    public static <T, K> Map<K, T> toMap(Collection<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, Function.identity()));
        }
        return new HashMap<>(0);
    }

    public static <T, K, V> Map<K, V> toMap(Collection<T> objects, Function<? super T, ? extends K> keyMapper,
                                            Function<? super T, ? extends V> valueMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, valueMapper));
        }
        return new HashMap<>(0);
    }

    public static <T, K> Map<K, T> toMapKey1(Collection<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, Function.identity(), (key1, key2) -> key1));
        }
        return new HashMap<>(0);
    }

    public static <T, K, V> Map<K, V> toMapKey1(Collection<T> objects, Function<? super T, ? extends K> keyMapper,
                                                Function<? super T, ? extends V> valueMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, valueMapper, (key1, key2) -> key1));
        }
        return new HashMap<>(0);
    }

    public static <T, K> Map<K, T> toMapKey2(Collection<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!(objects == null || objects.isEmpty())) {
            return objects.stream().collect(Collectors.toMap(keyMapper, Function.identity(), (key1, key2) -> key2));
        }
        return new HashMap<>(0);
    }

    public static <T, K> Map<K, List<T>> groupingBy(List<T> objects, Function<? super T, ? extends K> keyMapper) {
        if (!objects.isEmpty()) {
            return objects.stream().collect(Collectors.groupingBy(keyMapper));
        }
        return new HashMap<>(0);
    }

    /**
     * 对象转map
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
     * @param item 要转换的对象
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
     * @param obj1 取值对象
     * @param obj2 赋值对象
     */
    public static void objectSetValue(Object obj1, Object obj2){
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
     * @param obj1 取值对象
     * @param obj2 赋值对象
     * @param functions 要赋值的属性的get方法
     */
    @SafeVarargs
    public static <T, R> void objectSetValue(T obj1, Object obj2, TypeFunction<T, R>... functions){
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

    @SafeVarargs
    public static <T, R> void objectListSetValue(List<T> obj1, List<?> obj2, TypeFunction<T, R>... functions){
        if(CollectionUtils.isNotEmpty(obj1) && CollectionUtils.isNotEmpty(obj2) && obj2.size() >= obj1.size()) {
            Class<?> aClass = obj2.get(0).getClass();
            T t = obj1.get(0);
            for (TypeFunction<T, R> function : functions) {
                String setName = TypeFunction.getToSet(function);
                try {
                    Method declaredMethod = aClass.getDeclaredMethod(setName, function.apply(t).getClass());
                    for (int i = 0; i < obj1.size(); i++) {
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
}
