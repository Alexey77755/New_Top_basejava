package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.*;

public class MapUuidStorage extends AbstractStorage<String> {
    Map<String, Resume> hashmap = new HashMap<String, Resume>();

    @Override
    protected String getSearchKey(String uuid) {
            return uuid;
        }


    @Override
    protected boolean isExist(String uuid) {
        return hashmap.containsKey(uuid);
    }



    @Override
    protected void doSave(Resume r, String uuid) {

        hashmap.put(r.getUuid(), r);
    }

    @Override
    protected void doUpdate(Resume r, String uuid) {

        hashmap.put(uuid, r);
    }


    @Override
    protected void doDelete(String uuid) {
        hashmap.keySet().remove(uuid);
    }

    @Override
    protected Resume doGet(String uuid) {
        return hashmap.get(uuid);
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
