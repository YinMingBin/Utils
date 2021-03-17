package org.excel;

import com.sun.javafx.css.converters.StringConverter;
import jdk.nashorn.internal.runtime.regexp.RegExp;
import org.apache.commons.codec.StringEncoderComparator;
import org.custom.collection.MatchingCollectList;
import org.custom.collection.MatchingList;
import org.custom.function.VoidFunction;
import org.junit.Test;
import org.springframework.beans.propertyeditors.URIEditor;
import org.utils.BaseUtil;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.swing.text.StringContent;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;


/**
 * Unit test for simple App.
 */
public class AppTest{
    /**
     * Rigorous Test :-)
     */

    @Test
    public void test() throws UnsupportedEncodingException {
        System.out.println(URLEncoder.encode(" ", StandardCharsets.UTF_8.name()));
        System.out.println(numDistinct("anacondastreetracecar", "contra"));
    }

    private int sum = 0;
    private List<Map<Integer, Integer>> maps = new ArrayList<>();
    private Map<String, Integer> max = new HashMap<>();
    public int numDistinct(String s, String t) {
        int length = t.length();
        String[] stars = new String[length];
        for (int i = 0; i < length; i++) {
            String s1 = String.valueOf(t.charAt(i));
            stars[i] = s1;
            maps.add(new HashMap<>());
            max.put(s1, 0);
        }
        abc(s, stars, 0, 0);
        return sum;
    }

    public void abc(String s, String[] chars, int depth,int i){
        if (depth == chars.length){
            ++sum;
            return;
        }
        String aChar = chars[depth];
        int index = s.indexOf(aChar, i);
        Map<Integer, Integer> integerMap = null;
        Integer num = 0;
        String str = null;
        if (depth + 1 < chars.length) {
            Map<Integer, Integer> integerMap1 = maps.get(depth + 1);
            if (!integerMap1.isEmpty()) {
                integerMap = integerMap1;
                str = chars[depth + 1];
                num = integerMap.get(max.get(str));
            }

        }
        Map<Integer, Integer> integerMap1 = maps.get(depth);
        while (index >= 0){
            Integer integer = max.get(aChar);
            max.put(aChar, Math.max(integer, index));
            if (integerMap != null) {
                int index1 = s.indexOf(str, index + 1);
                Integer integer1 = integerMap.get(index1);
                if (integer1 != null) {
                    sum += Math.max(1, num - integer1);
                }
                return;
            }else {
                abc(s, chars, depth + 1, index + 1);
            }
            integerMap1.put(index, sum);
            index = s.indexOf(aChar, index + 1);
        }
    }
}
