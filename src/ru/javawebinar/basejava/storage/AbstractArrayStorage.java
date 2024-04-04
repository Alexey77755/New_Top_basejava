package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.exception.StorageException;
import ru.javawebinar.basejava.model.Resume;

import java.util.Arrays;
import java.util.List;


public abstract class AbstractArrayStorage extends AbstractStorage<Integer> {

    protected static final int STORAGE_LIMIT = 1000;
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
    protected void doSave(Resume r, Integer index) {
        if (size==STORAGE_LIMIT) {
            throw new StorageException("Storage overflow", r.getUuid());
        }
        else {
            insertElement( r,(Integer)index);
            size++;
        }
    }

    @Override
    protected void doDelete(Integer index) {
        fillDeletedElement((Integer)index);
        storage[size-1]=null;
        size--;
    }

    @Override
    protected Resume doGet(Integer index) {
        return storage[(Integer)index];
    }

    @Override
    protected  void doUpdate(Resume r, Integer index) {
        storage[(Integer)index]=r;
    }

   /* @Override
    public Resume[] getAll() {
        return Arrays.copyOfRange(storage, 0, size);
    }

    */

    @Override
    protected boolean isExist(Integer index) {
        return (Integer)index>=0;
    }

    @Override
    protected abstract Integer getSearchKey(String uuid) ;

    protected abstract void fillDeletedElement(int index);
    protected abstract void insertElement( Resume r,int index);

    @Override
    protected List<Resume> doCopyAll() {
        return Arrays.asList( Arrays.copyOfRange(storage,0,size));
    }
}
