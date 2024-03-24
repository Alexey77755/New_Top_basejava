package ru.javawebinar.basejava;

import ru.javawebinar.basejava.model.Resume;
import ru.javawebinar.basejava.storage.SortedArrayStorage;
import ru.javawebinar.basejava.storage.Storage;

/**
 * Test for your com.urise.webapp.storage.ArrayStorage implementation
 */
public class MainTestArrayStorage {
  private   static final Storage ARRAY_STORAGE = new SortedArrayStorage();

    public static void main(String[] args) {
        final  Resume r1;

            r1 = new Resume("uuid1");


        final  Resume r2;

            r2 = new Resume("uuid2");


        final  Resume r3;

            r3 = new Resume("uuid3");


        final   Resume r5;

            r5 = new Resume("uuid5");



        ARRAY_STORAGE.save(r1);
        ARRAY_STORAGE.save(r2);
        ARRAY_STORAGE.save(r3);
        ARRAY_STORAGE.save(r5);

        ARRAY_STORAGE.update(r5);
        System.out.println("uuid5: " + ARRAY_STORAGE.get("uuid5"));

        System.out.println("Get r1: " + ARRAY_STORAGE.get(r1.getUuid()));
        System.out.println("Size: " + ARRAY_STORAGE.size());

       // System.out.println("Get dummy: " + ARRAY_STORAGE.get("dummy"));

        printAll();
        ARRAY_STORAGE.delete(r1.getUuid());
        printAll();
        ARRAY_STORAGE.clear();
        printAll();

        System.out.println("Size: " + ARRAY_STORAGE.size());
    }

    static void printAll() {
        System.out.println("\nGet All");

        for (Resume r : ARRAY_STORAGE.getAllSorted()) {
            System.out.println(r);
        }
    }
}
