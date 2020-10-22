package com.main.Sign;

/**
 * ClassName:Students
 * Package:Sign
 * Description: 每次统计时的正常次数需要单独设置
 *
 * @date:2020/1/6 17:49
 * @author:Wang Jun
 */
public class Students {
    private String name;
    private int late = 0;   //迟到
    private int lose = 0; //缺卡
    private int evection = 0; //出差
    private int absenteeism = 0; //旷工
    private int normal = 26; //正常

    @Override
    public String toString() {
        return "Students{" +
                "name='" + name + '\'' +
                ", late=" + late +
                ", lose=" + lose +
                ", evection=" + evection +
                ", absenteeism=" + absenteeism +
                ", normal=" + normal +
                '}';
    }

    public int getAbsenteeism() {
        return absenteeism;
    }

    public void setAbsenteeism(int absenteeism) {
        this.absenteeism = absenteeism;
    }

    public Students() {
    }

    public Students(String name, int late, int lose, int evection, int absenteeism, int normal) {
        this.name = name;
        this.late = late;
        this.lose = lose;
        this.evection = evection;
        this.absenteeism = absenteeism;
        this.normal = normal;
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
        return (double) (this.normal+this.late)/26;
    }
}
