package org.excel;

import lombok.Data;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

@Data
public class User implements Serializable {
    private User descendant;
    private String name;
    private Integer age;
    private Integer sexId;
    private String sexName;
    private Integer statureId;
    private String stature;
    private Map<String, List<String>> map;
    private List<String> list;

    public void setSexName(String sexName){
        this.sexName = sexName;
    }

    public User(String name, Integer age) {
        this.name = name;
        this.age = age;
    }

    public User(String name, Integer age, Integer sexId) {
        this.name = name;
        this.age = age;
        this.sexId = sexId;
    }

    public User(String name, Integer age, Integer sexId, Integer statureId) {
        this.name = name;
        this.age = age;
        this.sexId = sexId;
        this.statureId = statureId;
    }
}

@Data
class Sex {
    private Integer sexId;
    private String sexName;
    private User user;

    public Sex(Integer sexId, String sexName) {
        this.sexId = sexId;
        this.sexName = sexName;
    }

    public Sex(Integer sexId, String sexName, User user) {
        this.sexId = sexId;
        this.sexName = sexName;
        this.user = user;
    }
}

@Data
class Stature {
    private Integer statureId;
    private String stature;

    public Stature(Integer statureId, String stature) {
        this.statureId = statureId;
        this.stature = stature;
    }
}
