package org.utils.roughly;

/**
 * 匹配对象赋值接口
 * @author Administrator
 * @param <T> 赋值对象类型
 * @param <I> 取值对象类型
 */
public interface MatchingInterface<T, I>{
    /**
     * 将i的值赋值给t
     * @param t 赋值对象
     * @param i 取值对象
     */
    void setValue(T t, I i);

}