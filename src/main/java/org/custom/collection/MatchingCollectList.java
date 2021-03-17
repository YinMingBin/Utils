package org.custom.collection;

import org.utils.roughly.MatchingCollectInterface;
import org.utils.definite.MatchingCollect;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 多个匹配赋值
 * @author Administrator
 */
public class MatchingCollectList<T> {
    private final List<MatchingCollectInterface<T>> matchingCollects = new ArrayList<>();

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
    public <I, R> MatchingCollectList<T> add(List<I> iList, Function<T, R> tFun, Function<I, R> iFun, MatchingList<T, I> matchingList){
        matchingCollects.add(new MatchingCollect<>(iList, tFun, iFun, matchingList));
        return this;
    }

    public void forEach(Consumer<? super MatchingCollectInterface<T>> action) {
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
