package org.example;

import org.junit.Test;


/**
 * Unit test for simple App.
 */
public class AppTest{
    /**
     * Rigorous Test :-)
     */
    @Test
    public void shouldAnswerWithTrue(){
        String[] strs = new String[10];
        Integer i = 0;
        System.out.println(strs.getClass().isArray());
        System.out.println(i.getClass().isArray());
    }
}
