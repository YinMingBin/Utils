package org.utils.object;

import org.custom.collection.MatchingList;
import org.custom.function.SetValueFunction;

import java.util.Map;
import java.util.function.Function;

/**
 * 同数据多个匹配赋值
 * @author Administrator
 */
public class MatchingKey<T, I, R>{
    private final Function<T, R> funT;
    private final MatchingList<T, I> matchingList;

    public MatchingKey(Function<T, R> funT) {
        this.funT = funT;
        this.matchingList = new MatchingList<>();
    }

    public <V> MatchingKey(Function<T, R> funT, SetValueFunction<T, V> assignFun, Function<I, V> valueFun) {
        this.funT = funT;
        this.matchingList = new MatchingList<>();
        this.matchingList.add(assignFun, valueFun);
    }

    public MatchingKey(Function<T, R> funT, MatchingList<T, I> matchingList) {
        this.funT = funT;
        this.matchingList = matchingList;
    }

    /**
     * 添加赋值对
     * @param assignFun 赋值函数
     * @param valueFun 取值函数
     * @param <V> 值类型
     */
    public <V> void add(SetValueFunction<T, V> assignFun, Function<I, V> valueFun){
        matchingList.add(assignFun, valueFun);
    }

    /**
     * 设置值
     * @param t 赋值对象
     * @param mapI 取值集合
     */
    public void setValue(T t, Map<?, I> mapI) {
        if (mapI != null && !mapI.isEmpty()) {
            matchingList.forEach(m -> m.setValue(t, mapI.get(funT.apply(t))));
        }
    }
}
