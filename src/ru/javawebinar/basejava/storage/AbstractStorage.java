package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.exception.ExistStorageException;
import ru.javawebinar.basejava.exception.NotExistStorageException;
import ru.javawebinar.basejava.exception.StorageException;
import ru.javawebinar.basejava.model.Resume;

public abstract class AbstractStorage implements Storage {

    @Override
    public void update(Resume r) throws CloneNotSupportedException {
        int index = getIndex(r.getUuid());
        if (index < 0) {
            throw new NotExistStorageException(r.getUuid());

        } else {
            updateElement(index, r);
        }
    }



    public void save(Resume r) throws CloneNotSupportedException {
        int index = getIndex(r.getUuid());
        if (index >= 0) {
            throw new ExistStorageException(r.getUuid());

        } else if (checkSize()) {
            throw new StorageException("Storage overflow", r.getUuid());

        } else {
            insertElement(size(), r);
            increaseSize();
        }
    }
    @Override
    public void delete(String uuid) throws CloneNotSupportedException {
        int index = getIndex(uuid);
        if (index < 0) {
            throw new NotExistStorageException(uuid);

        } else {
            fillDeletedElement(index,uuid);
            reduceSize();
        }
    }


    @Override
    final public Resume get(String uuid) throws CloneNotSupportedException {

        int index = getIndex(uuid);
        if (index < 0) {
            throw new NotExistStorageException(uuid);

        }
        return getResume(index, uuid) ;
    }

    protected abstract Resume getResume(int index,String uuid);

    protected abstract void updateElement(int index, Resume r) ;
    protected abstract int getIndex(String uuid) throws CloneNotSupportedException;

    protected abstract void insertElement(int index, Resume r);

    protected abstract boolean checkSize();

    protected abstract void fillDeletedElement(int index, String uuid);

    protected abstract void increaseSize();

    protected abstract void reduceSize();
}
