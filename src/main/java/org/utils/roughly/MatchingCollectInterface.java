package org.utils.roughly;

/**
 * 多个匹配赋值接口
 * @author Administrator
 * @param <T> 赋值对象类型
 */
public interface MatchingCollectInterface<T> {
    /**
     * 给对象赋值
     * @param t 赋值对象
     */
    void setValue(T t);
}