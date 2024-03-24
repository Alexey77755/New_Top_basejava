package ru.javawebinar.basejava;

import ru.javawebinar.basejava.model.Resume;
import ru.javawebinar.basejava.storage.MapResumeStorage;
import ru.javawebinar.basejava.storage.Storage;

import java.io.IOException;

/**
 * Interactive test for com.urise.webapp.storage.ArrayStorage implementation
 * (just run, no need to understand)
 */
public class MainMap {
   // private final static ArrayStorage ARRAY_STORAGE = new ArrayStorage();
    private static final String UUID_1 = "uuid1";
    private static final Resume RESUME_1 = new Resume(UUID_1);

    private static final String UUID_2 = "uuid2";
    private static final Resume RESUME_2 = new Resume(UUID_2);

    private static final String UUID_3 = "uuid3";
    private static final Resume RESUME_3 = new Resume(UUID_3);

    private static final String UUID_4 = "uuid4";
    private static final Resume RESUME_4 = new Resume(UUID_4);
    public static void main(String[] args) throws IOException, CloneNotSupportedException {
        Storage  collection = new MapResumeStorage();
        collection.save(RESUME_3);
        collection.save(RESUME_2);
        collection.save(RESUME_1);
        collection.getAllSorted();
    }



}
