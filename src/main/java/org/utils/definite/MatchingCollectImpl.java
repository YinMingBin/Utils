package org.utils.definite;

import lombok.Data;
import org.custom.collection.MatchingList;
import org.utils.BaseUtil;
import org.utils.roughly.MatchingCollect;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

/**
 * @author Administrator
 */
@Data
public class MatchingCollectImpl<T, I, R> implements MatchingCollect<T> {
    private Map<R, I> iMap;
    private Function<T, R> tFun;
    private Function<I, R> iFun;
    private MatchingList<T, I> matchingList;

    public MatchingCollectImpl(List<I> iList, Function<T, R> tFun, Function<I, R> iFun, MatchingList<T, I> matchingList) {
        this.iMap = iList != null ? BaseUtil.toMap(iList, iFun) : new HashMap<>(0);
        this.tFun = tFun;
        this.iFun = iFun;
        this.matchingList = matchingList;
    }

    @Override
    public void setValue(T t) {
        I i = iMap.get(tFun.apply(t));
        if (i != null) {
            matchingList.forEach(list -> list.setValue(t, i));
        }
    }
}