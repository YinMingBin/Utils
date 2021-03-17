package org.custom.collection;

import org.custom.function.SetValueFunction;
import org.utils.roughly.MatchingInterface;
import org.utils.definite.Matching;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 匹配赋值list
 * @author Administrator
 */
public class MatchingList<T, I> {
    private final List<MatchingInterface<T, I>> matchingList = new ArrayList<>();

    /**
     * 添加匹配赋值
     * @param assignFun 赋值函数
     * @param valueFun 取值函数
     * @param <R> 值类型
     * @return 本身
     */
    public <R> MatchingList<T, I> add(SetValueFunction<T, R> assignFun, Function<I, R> valueFun){
        matchingList.add(new Matching<>(assignFun, valueFun));
        return this;
    }

    public void forEach(Consumer<? super MatchingInterface<T, I>> action) {
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
