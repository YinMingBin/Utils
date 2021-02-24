package org.excel;

import jdk.nashorn.internal.runtime.regexp.RegExp;
import org.custom.collection.MatchingCollectList;
import org.custom.collection.MatchingList;
import org.junit.Test;
import org.utils.BaseUtil;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * Unit test for simple App.
 */
public class AppTest{
    /**
     * Rigorous Test :-)
     */

    @Test
    public void test() {
        String s = "-123";
        Pattern pattern = Pattern.compile("(?<=-)(?<b>\\d+)");
        Matcher matcher = pattern.matcher(s);
        matcher.find();
        System.out.println(matcher.group());
    }
    @Test
    public void shouldAnswerWithTrue(){
        List<Sex> sexList = new ArrayList<Sex>(){{
            add(new Sex(0, "女"));
            add(new Sex(1, "男"));
        }};
        List<Stature> statureList = new ArrayList<Stature>(){{
            add(new Stature(0, "胖"));
            add(new Stature(1, "瘦"));
        }};
        List<User> userList = new ArrayList<User>(){{
            add(new User("张三", 18, 1, 1));
            add(new User("李四", 16, 0, 0));
            add(new User("王五", 19, 1, 0));
        }};
        MatchingCollectList<User> matchingCollectList = new MatchingCollectList<>();
        matchingCollectList.add(sexList, User::getSexId, Sex::getSexId,
                new MatchingList<User, Sex>()
                        .add(User::setSexName, Sex::getSexName)
        ).add(statureList, User::getStatureId, Stature::getStatureId,
                new MatchingList<User, Stature>()
                        .add(User::setStature, Stature::getStature)
        );
        BaseUtil.matching(userList, matchingCollectList);
        System.out.println(123);
    }

}
