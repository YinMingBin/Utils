package org.utils.object;

import lombok.Data;
import org.custom.function.SetValueFunction;

import java.util.function.Function;

/**
 * 匹配对象赋值实现
 *
 * @param <T> 赋值对象类型
 * @param <I> 取值对象类型
 * @param <R> 值类型
 * @author Administrator
 */
@Data
public class Matching<T, I, R> {
    private SetValueFunction<T, R> assignFun;
    private Function<I, R> valueFun;

    public Matching(SetValueFunction<T, R> assignFun, Function<I, R> valueFun) {
        this.assignFun = assignFun;
        this.valueFun = valueFun;
    }

    public void setValue(T t, I i) {
        if (t != null && i != null) {
            assignFun.apply(t, valueFun.apply(i));
        }
    }
}