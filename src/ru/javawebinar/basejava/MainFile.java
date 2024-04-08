package ru.javawebinar.basejava;

import java.io.File;

public class MainFile {
    public static void fInput(String filePath) {
        File dir = new File(filePath);
        String[] list = dir.list();
        if (list != null) {
            for (String name : list) {
                if (new File(filePath + "\\" + name).isDirectory()) {
                    fInput(filePath + "\\" + name);
                }
                else {
                    System.out.println(name);
                }
            }
        }
    }

    public static void main(String[] args) {
        String filePath = "D:\\Educational_project\\basejava";
       // File file = new File("./.gitignore");
       /* try {
            System.out.println(file.getCanonicalPath());
        } catch (IOException e) {
            throw new RuntimeException("Error", e);
        }*/
        //File dir = new File("D:\\Educational_project\\basejava\\src\\ru\\javawebinar\\basejava");
        //System.out.println(dir.isDirectory());
        //System.out.println(dir.listFiles());
        fInput(filePath);
       /* try (FileInputStream fis = new FileInputStream(filePath)) {

            System.out.println(fis.read());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }*/
    }
}
