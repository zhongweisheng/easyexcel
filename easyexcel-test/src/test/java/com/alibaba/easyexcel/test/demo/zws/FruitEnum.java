package com.alibaba.easyexcel.test.demo.zws;

public enum FruitEnum {
    APPLE("苹果"),
    BANANA("香蕉"),
    ORANGE("橙子"),
    WATERMELON("西瓜");

    private String name;

    FruitEnum(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }
}
