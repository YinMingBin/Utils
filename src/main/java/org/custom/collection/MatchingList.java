package org.custom.collection;

import org.custom.function.SetValue;
import org.utils.roughly.Matching;
import org.utils.definite.MatchingImpl;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 匹配赋值list
 * @author Administrator
 */
public class MatchingList<T, I> {
    private final List<Matching<T, I>> matchingList = new ArrayList<>();

    /**
     * 添加匹配赋值
     * @param assignFun 赋值函数
     * @param valueFun 取值函数
     * @param <R> 值类型
     * @return 本身
     */
    public <R> MatchingList<T, I> add(SetValue<T, R> assignFun, Function<I, R> valueFun){
        matchingList.add(new MatchingImpl<>(assignFun, valueFun));
        return this;
    }

    public void forEach(Consumer<? super Matching<T, I>> action) {
        matchingList.forEach(action);
    }

    /**
     * 判断是否有值
     * @return 是否有值
     */
    public boolean isNotEmpty() {
        return !matchingList.isEmpty();
    }
}
