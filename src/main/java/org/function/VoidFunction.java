package org.function;

/**
 * 无返回值函数
 * @author Administrator
 */
public interface VoidFunction {
    interface One<V> {
        /**
         * 一个参数
         * @param val 参数
         */
        void apply(V val);
    }

    interface Two<V1, V2> {
        /**
         * 两个参数
         * @param val1 参数1
         * @param val2 参数2
         */
        void apply(V1 val1, V2 val2);
    }
}