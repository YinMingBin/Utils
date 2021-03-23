package org.custom.function;

import java.beans.Introspector;
import java.io.Serializable;
import java.lang.invoke.SerializedLambda;
import java.lang.reflect.Method;
import java.util.function.Function;


/**
 * Function
 *
 * @author Administrator
 */
@FunctionalInterface
public interface TypeFunction<T, R> extends Serializable, Function<T, R> {

    /**
     * 获取列名称
     *
     * @param lambda ::
     * @return String
     */
    static String getLambdaColumnName(Serializable lambda) {
        try {
            Method method = lambda.getClass().getDeclaredMethod("writeReplace");
            method.setAccessible(Boolean.TRUE);
            SerializedLambda serializedLambda = (SerializedLambda) method.invoke(lambda);
            String getter = serializedLambda.getImplMethodName();
            return Introspector.decapitalize(getter.replace("get", ""));
        } catch (ReflectiveOperationException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * get方法获取set方法名
     *
     * @param lambda ::
     * @return String
     */
    static String getToSet(Serializable lambda) {
        try {
            Method method = lambda.getClass().getDeclaredMethod("writeReplace");
            method.setAccessible(Boolean.TRUE);
            SerializedLambda serializedLambda = (SerializedLambda) method.invoke(lambda);
            String getter = serializedLambda.getImplMethodName();
            return getter.replace("get", "set");
        } catch (ReflectiveOperationException e) {
            throw new RuntimeException(e);
        }
    }

}