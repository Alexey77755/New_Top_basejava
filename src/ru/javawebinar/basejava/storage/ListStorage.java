package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.ArrayList;

public class ListStorage extends AbstractStorage  {
    ArrayList<Resume> arr = new ArrayList<Resume>();

    @Override
    public void clear() {
        arr.clear();
    }

    @Override
    public Resume[] getAll() {
        Resume[] array = new Resume[arr.size()];
        return arr.toArray(array);

    }

    @Override
    protected Resume getResume(int index,String uuid) {
        return arr.get(index);
    }

    @Override
    protected void updateElement(int index, Resume r) {
        arr.set(index, r);
    }


    @Override
    public int size() {
        return arr.size();
    }

    public int getIndex(String uuid) throws CloneNotSupportedException {

            return arr.indexOf(new Resume(uuid));


    }

    @Override
    protected void insertElement(int index, Resume r) {
        arr.add(index, r);
    }



    @Override
    protected boolean checkSize() {
        return false;
    }

    @Override
    protected void fillDeletedElement(int index, String uuid) {
        arr.remove(index);
    }

    @Override
    protected void increaseSize() {

    }

    @Override
    protected void reduceSize() {

    }


}
