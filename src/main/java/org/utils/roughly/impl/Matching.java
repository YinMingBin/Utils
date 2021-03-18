package org.utils.roughly.impl;

import lombok.Data;
import org.custom.function.SetValueFunction;
import org.utils.roughly.MatchingInterface;

import java.util.function.Function;

/**
 * 匹配对象赋值实现
 * @author Administrator
 * @param <T> 赋值对象类型
 * @param <I> 取值对象类型
 * @param <R> 值类型
 */
@Data
public class Matching<T, I, R> implements MatchingInterface<T, I> {
    private SetValueFunction<T, R> assignFun;
    private Function<I, R> valueFun;

    public Matching(SetValueFunction<T, R> assignFun, Function<I, R> valueFun) {
        this.assignFun = assignFun;
        this.valueFun = valueFun;
    }

    @Override
    public void setValue(T t, I i) {
        assignFun.apply(t, valueFun.apply(i));
    }
}