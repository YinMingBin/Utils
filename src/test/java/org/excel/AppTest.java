package org.excel;

import lombok.Data;
import org.apache.poi.ss.formula.functions.T;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;
import org.custom.function.NoParamsFunction;
import org.custom.function.VoidFunction;
import org.junit.Test;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.UnsupportedEncodingException;
import java.lang.annotation.*;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * Unit test for simple App.
 */
public class AppTest {
    /**
     * Rigorous Test :-)
     */

    @Test
    public void test() throws UnsupportedEncodingException {
        str(null);
    }

    @Default()
    public static void str(@StringDefault("aaa") String str){
        System.out.println(str);
    }

    @Target(ElementType.METHOD)
    @Retention(RetentionPolicy.RUNTIME)
    @Documented
    @interface Default{
    }

    @Target(ElementType.PARAMETER)
    @Retention(RetentionPolicy.RUNTIME)
    @Documented
    @interface StringDefault{
        String value() default "";
    }

    @Component
    @Aspect
    class DefaultAop{

        @Around("@annotation(Default)")
        public Object setDefault(ProceedingJoinPoint joinPoint){
            final Object[] args = joinPoint.getArgs();
            return null;
        }
    }
}

