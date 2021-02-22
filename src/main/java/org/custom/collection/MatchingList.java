package org.custom.collection;

import org.custom.function.SetValue;
import org.utils.BaseUtil;

import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 匹配赋值list
 * @author Administrator
 */
public class MatchingList<T, I> {
    private final List<BaseUtil.Matching<T, I>> matchingList = new ArrayList<>();

    /**
     * 添加匹配赋值
     * @param assignFun 赋值函数
     * @param valueFun 取值函数
     * @param <R> 值类型
     * @return 本身
     */
    public <R> MatchingList<T, I> add(SetValue<T, R> assignFun, Function<I, R> valueFun){
        matchingList.add(BaseUtil.matchingProperty(assignFun, valueFun));
        return this;
    }

    public void forEach(Consumer<? super BaseUtil.Matching<T, I>> action) {
        matchingList.forEach(action);
    }
}
