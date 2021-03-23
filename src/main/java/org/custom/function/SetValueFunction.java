package org.custom.function;

/**
 * 对象属性赋值
 *
 * @param <T> 对象类型
 * @param <V> 值类型
 * @author Administrator
 */
@FunctionalInterface
public interface SetValueFunction<T, V> {

    /**
     * 给对象属性赋值
     *
     * @param t 对象
     * @param v 值
     */
    void apply(T t, V v);
}