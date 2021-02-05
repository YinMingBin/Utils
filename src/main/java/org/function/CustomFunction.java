package org.function;

/**
 * 自定义函数
 * @author Administrator
 */
public interface CustomFunction {

    /**
     * 无参有反函数接口
     * @param <R>
     */
    @FunctionalInterface
    interface NoParams<R> {
        /**
         * 无参
         * @return R
         */
        R apply();
    }
}
