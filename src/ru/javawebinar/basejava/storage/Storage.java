package ru.javawebinar.basejava.storage;

import  ru.javawebinar.basejava.model.Resume;

/**
 * Array based storage for Resumes
 */

public interface Storage {


     void clear();

     void update(Resume r) throws CloneNotSupportedException;

     void save(Resume r) throws CloneNotSupportedException;

     Resume get(String uuid) throws CloneNotSupportedException;

     void delete(String uuid) throws CloneNotSupportedException;

    /**
     * @return array, contains only Resumes in storage (without null)
     */
     Resume[] getAll();

     int size();

}
