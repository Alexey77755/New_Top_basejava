package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.Arrays;

public class SortedArrayStorage extends AbstractArrayStorage{


    @Override
    public void clear() {
        Arrays.fill(storage, 0, size, null);
        size = 0;
    }

    @Override
    public void update(Resume r) {
        int index = getIndex(r.getUuid());
        if (index < 0) {
            System.out.println("Resume "  + r.getUuid() + " not exist");

        } else {
            storage[index] = r;
        }
    }

    @Override
    public void save(Resume r) {
        if (getIndex(r.getUuid()) > 0) {
            System.out.println("Resume " + r.getUuid() + " already exist");

        } else if (size == STORAGE_LIMIT) {
            System.out.println("Storage overflow");

        } else {
            storage[-(Arrays.binarySearch(storage,0,size, r)+1)] = r;
            size++;
        }
    }

    @Override
    public void delete(String uuid) {
        int index = getIndex(uuid);
        if (index < 0) {
            System.out.println("Resume " + uuid + " not exist");

        } else {
            System.arraycopy(storage,index+1,storage,index,size - 1);
            size--;
        }
    }

    @Override
    public Resume[] getAll() {
        return Arrays.copyOfRange(storage, 0, size);
    }
    @Override
    protected int getIndex(String uuid) {
        Resume searchKey =new Resume();
        searchKey.setUuid(uuid);
        return Arrays.binarySearch(storage,0,size, searchKey);
    }
}
