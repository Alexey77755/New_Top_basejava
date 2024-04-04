package ru.javawebinar.basejava;

import java.io.File;
import java.io.IOException;

public class MainFile {
    public static void main(String[] args) throws IOException {
        File file = new File("./.gitignore");
        System.out.println(file.getCanonicalPath());
        File dir = new File("D:\\Educational_project\\basejava\\src\\ru\\javawebinar\\basejava");
        System.out.println(dir.isDirectory());
        System.out.println(dir.listFiles());
        String[] list =dir.list();
                if(list!=null) {
                    for (String name :list) {
                        System.out.println(name);
                    }
                }
    }
}
