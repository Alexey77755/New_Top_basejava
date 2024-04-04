package ru.javawebinar.basejava;


import ru.javawebinar.basejava.model.SectionType;

public class TestSingleton {
    private static TestSingleton instance ;
    public static TestSingleton instance(){
        if(instance==null){
            instance= new TestSingleton();
        }
        return  instance;}
    private TestSingleton(){

    }

    public static void main(String[] args) {
        TestSingleton.instance().toString();
        Singleton instance = Singleton.valueOf("INSTANCE");
        System.out.println(instance.ordinal());
       for(SectionType type : SectionType.values()){
            System.out.println(type);
        }
    }
    public  enum Singleton {
        INSTANCE
    }
}
