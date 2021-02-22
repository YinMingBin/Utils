package org.excel;

import org.custom.collection.MatchingList;
import org.junit.Test;
import org.utils.BaseUtil;

import java.util.*;


/**
 * Unit test for simple App.
 */
public class AppTest{
    /**
     * Rigorous Test :-)
     */
    @Test
    public void shouldAnswerWithTrue(){
        List<Sex> sexList = new ArrayList<Sex>(){{
            add(new Sex(0, "女"));
            add(new Sex(1, "男"));
        }};
        List<User> userList = new ArrayList<User>(){{
            add(new User("张三", 18, 1));
            add(new User("李四", 16, 0));
            add(new User("王五", 19, 1));
        }};
        BaseUtil.matching(userList, sexList, User::getSexId, Sex::getSexId,
                new MatchingList<User, Sex>()
                        .add(User::setSexName, Sex::getSexName)
        );
        System.out.println(123);
    }

    public String convert(String s, int numRows) {
        StringBuilder[] sbs = new StringBuilder[numRows];
        int len = s.length();
        int i = 0;
        int j = 0;
        boolean b = true;
        while (i < len){
            StringBuilder sb = sbs[j];
            if(sb == null){
                sb = new StringBuilder();
                sbs[j] = sb;
            }
            sb.append(s.charAt(i));
            i++;
            j = b ? ++j : --j;
            if(numRows < 2){j = 0;}
            if(b){
                b = j < numRows - 1;
            }else {
                b = j == 0;
            }
        }
        StringBuilder sb = sbs[0];
        for (int i1 = 1; i1 < numRows; i1++) {
            StringBuilder sb1 = sbs[i1];
            if(sb1 != null) {
                sb.append(sb1.toString());
            }
        }
        return sb.toString();
    }
}
