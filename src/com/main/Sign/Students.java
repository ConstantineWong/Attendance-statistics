package com.main.Sign;

/**
 * ClassName:Students
 * Package:Sign
 * Description:
 *
 * @date:2020/1/6 17:49
 * @author:Wang Jun
 */
public class Students {
    private String name = "未知";
    private int late = 0;   //迟到
    private int lose = 0; //缺卡
    private int evection = 0; //出差
    private int normal = 26; //正常

    @Override
    public String toString() {
        return "Students{" +
                "name='" + name + '\'' +
                ", late=" + late +
                ", lose=" + lose +
                ", evection=" + evection +
                ", normal=" + normal +
                '}';
    }

    public String getName() {
        return name;
    }

    public int getNormal() {
        return normal;
    }

    public void setNormal(int normal) {
        this.normal = normal;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getLate() {
        return late;
    }

    public void setLate(int late) {
        this.late = late;
    }

    public int getLose() {
        return lose;
    }

    public void setLose(int lose) {
        this.lose = lose;
    }

    public int getEvection() {
        return evection;
    }

    public void setEvection(int evection) {
        this.evection = evection;
    }

    public double attendrat(){
        return (double) this.normal/26;
    }
}
