package com.urise.webapp.storage;

import com.urise.webapp.model.Resume;

import java.util.Arrays;

/**
 * Array based storage for Resumes
 */

public class ArrayStorage {
    private Resume[] storage = new Resume[10000];
    private int size;

    public void clear() {
        Arrays.fill(storage, null);
        size = 0;
    }

    public void update(Resume r) {
        if (exist_Resume(r) != 0) {
            storage[exist_Resume(r)] = r;
        } else {
            System.out.println("Resume is not exist");
        }
    }

    public void save(Resume r) {
        if (exist_Resume(r) == 0) {
            size = size + 1;
            storage[size] = r;
        } else {
            System.out.println("Resume is exist");
        }

    }

    public Resume get(String uuid) {
        Resume r = new Resume();
        r.setUuid(uuid);
        if (exist_Resume(r) != 0) {
                    return storage[exist_Resume(r)];
        }
       else {
                System.out.println("Resume is not exist");
       }
        return null;
    }

    public void delete(String uuid) {
       Resume r = new Resume();
       r.setUuid(uuid);
        if (exist_Resume(r) != 0) {
            for (int i = 0; i < size; i++) {
                if (storage[i].getUuid() == uuid) {
                    /* System.arraycopy(storage, i + 1, storage, i, storage.length - 1);*/
                    storage[i] = storage[size - 1];
                    storage[size - 1] = null;
                    size--;
                }

            }
        } else {
            System.out.println("Resume is not exist");
        }
    }

    /**
     * @return array, contains only Resumes in storage (without null)
     */
    public Resume[] getAll() {
        return Arrays.copyOfRange(storage, 0, size);
    }

    public int size() {
        return size;
    }

    private int exist_Resume(Resume r) {
        for (int i = 0; i < size; i++) {
            if (storage[i].getUuid() == r.getUuid()) {
                return i;
            }
        }
        return 0;
    }

}
