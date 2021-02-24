package org.custom.function;

/**
 * 对象属性赋值
 * @author Administrator
 * @param <T> 对象类型
 * @param <R> 值类型
 */
@FunctionalInterface
public interface SetValueFunction<T, R> {

    /**
     * 给对象属性赋值
     * @param t 对象
     * @param r 值
     */
    void apply(T t, R r);
}