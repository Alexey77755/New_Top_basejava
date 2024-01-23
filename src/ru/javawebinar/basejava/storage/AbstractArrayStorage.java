package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.Arrays;


public abstract class AbstractArrayStorage extends AbstractStorage {

    protected static final int STORAGE_LIMIT = 10000;
    protected Resume[] storage = new Resume[STORAGE_LIMIT];
    protected int size = 0;

    @Override
    public int size() {
        return size;
    }
    public void clear() {
        Arrays.fill(storage, 0, size, null);
        size = 0;
    }
    @Override
    public Resume[] getAll() {
        return Arrays.copyOfRange(storage, 0, size);
    }


    @Override
    public boolean  checkSize(){
       if (size== STORAGE_LIMIT) {
           return true;
       }
       return false;
    }
    @Override
    public void increaseSize(){
       size++;
    }
    @Override
    public void  reduceSize(){
        storage[size - 1] = null;
        size--;
    }
    public Resume getResume(int index){
       return storage[index];
    }

    @Override
    protected void updateElement(int index, Resume r) {
        storage[index]=r;
    }


    //protected abstract void insertElement(Resume r, int index);

}
