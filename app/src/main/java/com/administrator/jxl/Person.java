package com.administrator.jxl;

public class Person {
    private String age;
    private String name;
    private String sex;
  public Person(String age,String name,String sex){
      this.age=age;
      this.name=name;
      this.sex=sex;

  }
    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getName() {
        return name;
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
}
