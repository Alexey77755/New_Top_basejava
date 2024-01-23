package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

/**
 * Array based storage for Resumes
 */

public class ArrayStorage extends AbstractArrayStorage {


    protected int getIndex(String uuid) {
        if (size() != 0) {
            for (int i = 0; i < size; i++) {
                if (uuid.equals(storage[i].getUuid())) {
                    return i;
                }
            }
        }
        return -1;
    }
    @Override
    protected void fillDeletedElement(int index) {

        storage[index] = storage[size - 1];
    }

    @Override
    protected void insertElement( int index,Resume r) {
        storage[size] = r;
    }

}
