package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.Arrays;

public class SortedArrayStorage extends AbstractArrayStorage{

       @Override
    protected int getIndex(String uuid) {
           Resume searchKey = null;
           try {
               searchKey = new Resume(uuid);
           } catch (CloneNotSupportedException e) {
               throw new RuntimeException(e);
           }
           return Arrays.binarySearch(storage,0,size, searchKey);
    }

    @Override
    protected void fillDeletedElement(int index, String uuid) {
           int numMoved=size-index-1;
        System.arraycopy(storage,index+1,storage,index,numMoved);
    }

    @Override
    protected void insertElement (int index,Resume r) {
        int insertIdx=-index-1;
        System.arraycopy(storage,insertIdx,storage,insertIdx+1,size-insertIdx);
        storage[insertIdx] =r;
    }
}
