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

    public <V> void add(SetValueFunction<T, V> assignFun, Function<I, V> valueFun){
        matchingList.add(assignFun, valueFun);
    }

    public void setValue(T t, Map<?, I> mapI) {
        matchingList.forEach(m -> m.setValue(t, mapI.get(funT.apply(t))));
    }
}
