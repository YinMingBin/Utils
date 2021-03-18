package org.utils.object;

import lombok.Data;
import org.custom.collection.MatchingList;
import org.custom.function.SetValueFunction;
import org.utils.BaseUtil;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;


/**
 * 匹配赋值
 * @author Administrator
 * @param <T> 赋值对象类型
 * @param <I> 取值对象类型
 * @param <R> 匹配字段类型
 */
@Data
public class MatchingCollect<T, I, R>{
    private Map<R, I> iMap;
    private Function<T, R> tFun;
    private Function<I, R> iFun;
    private MatchingList<T, I> matchingList;

    public MatchingCollect(List<I> iList, Function<T, R> tFun, Function<I, R> iFun, MatchingList<T, I> matchingList) {
        this.iMap = iList != null ? BaseUtil.toMapKey1(iList, iFun) : new HashMap<>(0);
        this.tFun = tFun;
        this.iFun = iFun;
        this.matchingList = matchingList;
    }

    public MatchingCollect(List<I> iList, Function<T, R> tFun, Function<I, R> iFun,
                           SetValueFunction<T, R> assignFun, Function<I, R> valueFun) {
        this.iMap = iList != null ? BaseUtil.toMapKey1(iList, iFun) : new HashMap<>(0);
        this.tFun = tFun;
        this.iFun = iFun;
        this.matchingList = new MatchingList<>();
        this.matchingList.add(assignFun, valueFun);
    }

    public void setValue(T t) {
        I i = iMap.get(tFun.apply(t));
        if (i != null) {
            matchingList.forEach(list -> list.setValue(t, i));
        }
    }
}