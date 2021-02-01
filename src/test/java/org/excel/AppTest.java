package org.excel;

import org.apache.commons.collections4.CollectionUtils;
import org.junit.Test;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;


/**
 * Unit test for simple App.
 */
public class AppTest{
    /**
     * Rigorous Test :-)
     */
    @Test
    public void shouldAnswerWithTrue() throws IOException, ClassNotFoundException {
        List<Solution> solutionList = new ArrayList<>(2);
        Solution solution = new Solution("张三", 18, "男");
        solutionList.add(solution);
        solutionList.add(new Solution("李四", 19, "女"));
        List<Solution1> solutionList1 = new ArrayList<>(2);
        Solution1 solution1 = new Solution1();
        solutionList1.add(solution1);
        solutionList1.add(new Solution1());
        objectListSetValue(solutionList, solutionList1, Solution::getName, Solution::getAge, Solution::getSex);
        System.out.println(solutionList.toString());
        System.out.println(solutionList1.toString());

    }

    /**
     * 将obj1的值赋给obj2（获取get方法将get改成set获取方法调用）
     * @param obj1 取值对象
     * @param obj2 赋值对象
     */
    public static void objectSetValue(Object obj1, Object obj2){
        try {
            Class<?> aClass = obj1.getClass();
            Class<?> aClass1 = obj2.getClass();
            for (Method method : aClass.getDeclaredMethods()) {
                String name = method.getName();
                if(name.startsWith("get")) {
                    Object invoke = method.invoke(obj1);
                    String replace = name.replace("get", "set");
                    try {
                        Method declaredMethod = aClass1.getDeclaredMethod(replace, method.getReturnType());
                        declaredMethod.invoke(obj2, invoke);
                    } catch (NoSuchMethodException e) {
                        System.out.println("赋值失败：没有" + replace + "方法");
                    }
                }
            }
        } catch (IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将obj1的值赋给obj2（直接操作属性）（不安全）
     * @param obj1 取值对象
     * @param obj2 赋值对象
     */
    public static void objectSetValue1(Object obj1, Object obj2){
        Class<?> aClass = obj1.getClass();
        Class<?> aClass1 = obj2.getClass();
        Field[] fields = aClass.getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            String name = field.getName();
            try {
                Field declaredField = aClass1.getDeclaredField(name);
                declaredField.setAccessible(true);
                declaredField.set(obj2, field.get(obj1));
            } catch (NoSuchFieldException e) {
                System.out.println("没有" + name + "属性");
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }
    /**
     * 将obj1的值赋给obj2（通过属性名调用get/set方法）
     * @param obj1 取值对象
     * @param obj2 赋值对象
     */
    public static void objectSetValue2(Object obj1, Object obj2){
        Class<?> aClass = obj1.getClass();
        Class<?> aClass1 = obj2.getClass();
        Field[] fields = aClass.getDeclaredFields();
        for (Field field : fields) {
            String name = field.getName();
            char[] chars = name.toCharArray();
            chars[0] -= 32;
            String name2 = String.valueOf(chars);
            try {
                Method declaredMethod1 = aClass.getDeclaredMethod("get" + name2);
                Object invoke = declaredMethod1.invoke(obj1);
                Method declaredMethod = aClass1.getDeclaredMethod("set" + name2, declaredMethod1.getReturnType());
                declaredMethod.invoke(obj2, invoke);
            } catch (IllegalAccessException | InvocationTargetException e) {
                System.out.println("属性" + name + "的方法调用失败!!!");
                e.printStackTrace();
            } catch (NoSuchMethodException e) {
                System.out.println("属性" + name + "没有get/set方法!!!");
                e.printStackTrace();
            }
        }
    }

    @SafeVarargs
    public static <T, R> void objectSetValue(T obj1, Object obj2, TypeFunction<T, R>... functions){
        Class<?> aClass = obj2.getClass();
        for (TypeFunction<T, R> function : functions) {
            R apply = function.apply(obj1);
            String setName = TypeFunction.getToSet(function);
            try {
                Method declaredMethod = aClass.getDeclaredMethod(setName, apply.getClass());
                declaredMethod.invoke(obj2, apply);
            } catch (NoSuchMethodException e) {
                System.out.println("没有" + setName + "方法");
                e.printStackTrace();
            } catch (IllegalAccessException | InvocationTargetException e) {
                e.printStackTrace();
            }
        }
    }

    @SafeVarargs
    public static <T, R> void objectListSetValue(List<T> obj1, List<?> obj2, TypeFunction<T, R>... functions){
        if(CollectionUtils.isNotEmpty(obj1) && CollectionUtils.isNotEmpty(obj2) && obj2.size() >= obj1.size()) {
            Class<?> aClass = obj2.get(0).getClass();
            T t = obj1.get(0);
            for (TypeFunction<T, R> function : functions) {
                String setName = TypeFunction.getToSet(function);
                try {
                    Method declaredMethod = aClass.getDeclaredMethod(setName, function.apply(t).getClass());
                    for (int i = 0; i < obj1.size(); i++) {
                        R apply = function.apply(obj1.get(i));
                        declaredMethod.invoke(obj2.get(i), apply);
                    }
                } catch (NoSuchMethodException e) {
                    System.out.println("没有" + setName + "方法");
                    e.printStackTrace();
                } catch (IllegalAccessException | InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
    }


    static class Solution implements Serializable{
        private String name;
        private Integer age;
        private String sex;

        public String getName() {
            return name;
        }

        public String getAbc() {
            return name;
        }

        public Integer getAge() {
            return age;
        }

        public void setName(String name) {
            this.name = name;
        }

        public void setAge(Integer age) {
            this.age = age;
        }

        public String getSex() {
            return sex;
        }

        public void setSex(String sex) {
            this.sex = sex;
        }

        public Solution() {
        }

        public Solution(String name, Integer age, String sex) {
            this.name = name;
            this.age = age;
            this.sex = sex;
        }

        @Override
        public String toString() {
            return "Solution1{" +
                    "name='" + name + '\'' +
                    ", age=" + age +
                    ", sex='" + sex + '\'' +
                    '}';
        }
    }

    static class Solution1 {
        private String name;
        private Integer age;
        private String sex;

        public String getName() {
            return name;
        }

        public Integer getAge() {
            return age;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getSex() {
            return sex;
        }

        public void setSex(String sex) {
            this.sex = sex;
        }

        public Solution1() {
        }

        public Solution1(String name, Integer age) {
            this.name = name;
            this.age = age;
        }

        @Override
        public String toString() {
            return "Solution1{" +
                    "name='" + name + '\'' +
                    ", age=" + age +
                    ", sex='" + sex + '\'' +
                    '}';
        }
    }

    public List<List<Integer>> generate(int numRows) {
        List<List<Integer>> lists = new ArrayList<>(numRows);
        if(numRows>0) {
            List<Integer> l = new ArrayList<>(1);
            l.add(1);
            lists.add(l);
            for (int i = 1; i < numRows; i++) {
                List<Integer> list2 = lists.get(i - 1);
                List<Integer> list = new ArrayList<>(i + 1);
                for (int j = 0; j < i + 1; j++) {
                    if (j > 0 && j < i) {
                        list.add(list2.get(j - 1) + list2.get(j));
                    } else {
                        list.add(1);
                    }
                }
                System.out.println(list);
                lists.add(list);
            }
        }
        return lists;
    }

    enum Task {
        IQC
    }

    public static <T> T copy(T obj){
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            new ObjectOutputStream(baos).writeObject(obj);
            ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
            Object o = new ObjectInputStream(bais).readObject();
            return (T) o;
        } catch (IOException | ClassNotFoundException e) {
            e.printStackTrace();
        }
        return null;
    }

}
