package org.custom.function;

/**
 * 无参有反函数接口
 * @author Administrator
 * @param <R>
 */
@FunctionalInterface
public interface NoParamsFunction<R> {
    /**
     * 无参有反
     * @return R
     */
    R apply();
}