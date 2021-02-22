package org.excel;

import lombok.Data;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

@Data
public class User implements Serializable {
    private String name;
    private Integer age;
    private Integer sexId;
    private String sexName;
    private Map<String, List<String>> map;
    private List<String> list;

    public User(String name, Integer age) {
        this.name = name;
        this.age = age;
    }

    public User(String name, Integer age, Integer sexId) {
        this.name = name;
        this.age = age;
        this.sexId = sexId;
    }
}

@Data
class Sex {
    private Integer sexId;
    private String sexName;

    public Sex(Integer sexId, String sexName) {
        this.sexId = sexId;
        this.sexName = sexName;
    }
}
