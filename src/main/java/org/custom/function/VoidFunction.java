package org.custom.function;

/**
 * 无返回值函数
 * @author Administrator
 */
@FunctionalInterface
public interface VoidFunction {

    /**
     * 无参数
     */
    void apply();

    @FunctionalInterface
    interface One<V> {
        /**
         * 一个参数
         * @param val 参数
         */
        void apply(V val);
    }

    @FunctionalInterface
    interface Two<V1, V2> {
        /**
         * 两个参数
         * @param val1 参数1
         * @param val2 参数2
         */
        void apply(V1 val1, V2 val2);
    }
}