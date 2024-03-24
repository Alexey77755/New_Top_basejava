package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.*;

public class MapUuidStorage extends AbstractStorage {
    Map<String, Resume> hashmap = new HashMap<String, Resume>();

    @Override
    protected String getSearchKey(String uuid) {
            return uuid;
        }


    @Override
    protected boolean isExist(Object uuid) {
        return hashmap.containsKey((String) uuid);
    }



    @Override
    protected void doSave(Resume r, Object uuid) {

        hashmap.put(r.getUuid(), r);
    }

    @Override
    protected void doUpdate(Resume r, Object uuid) {

        hashmap.put((String) uuid, r);
    }


    @Override
    protected void doDelete(Object uuid) {
        hashmap.keySet().remove((String) uuid);
    }

    @Override
    protected Resume doGet(Object uuid) {
        return hashmap.get((String) uuid);
    }

    @Override
    protected List<Resume> doCopyAll() {

        return new ArrayList<>( hashmap.values()) ;
    }

    @Override
    public void clear() {
        hashmap.clear();
    }



    /* @Override
     public Resume[] getAll() {
         return hashmap.values().toArray(new Resume[hashmap.size()] );
     }
 */


    @Override
    public int size() {
        return hashmap.size();
    }


}
