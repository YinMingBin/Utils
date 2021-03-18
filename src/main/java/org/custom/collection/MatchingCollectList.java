package org.custom.collection;

import org.utils.object.MatchingCollect;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 多个匹配赋值
 * @author Administrator
 */
public class MatchingCollectList<T, I> {
    private final List<MatchingCollect<T, I, ?>> matchingCollects = new ArrayList<>();

    /**
     * 添加匹配赋值
     * @param iList 取值对象集
     * @param tFun 赋值对象的匹配字段函数
     * @param iFun 取值对象的匹配字段函数
     * @param matchingList 要赋值的字段
     * @param <I> 取值对象类型
     * @param <R> 匹配字段类型
     * @return 本身
     */
    public <R> MatchingCollectList<T, I> add(List<I> iList, Function<T, R> tFun, Function<I, R> iFun, MatchingList<T, I> matchingList){
        matchingCollects.add(new MatchingCollect<T, I, R>(iList, tFun, iFun, matchingList));
        return this;
    }

    public void forEach(Consumer<? super MatchingCollect<T, I, ?>> action) {
        matchingCollects.forEach(action);
    }

    /**
     * 判断是否有值
     * @return 是否有值
     */
    public boolean isNotEmpty() {
        return !matchingCollects.isEmpty();
    }
}
