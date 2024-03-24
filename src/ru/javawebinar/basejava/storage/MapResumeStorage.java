package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MapResumeStorage extends AbstractStorage {
    Map<String, Resume> hashmap = new HashMap<String, Resume>();

    @Override
    protected Resume getSearchKey(String uuid) {

        return hashmap.get(uuid);
        }


    @Override
    protected boolean isExist(Object resume) {
        return resume!=null;
    }



    @Override
    protected void doSave(Resume r, Object resume) {

        hashmap.put(r.getUuid(), r);
    }

    @Override
    protected void doUpdate(Resume r, Object resume) {

        hashmap.put(r.getUuid(), r);
    }


    @Override
    protected void doDelete(Object resume) {
        hashmap.keySet().remove(((Resume)resume).getUuid());
    }

    @Override
    protected Resume doGet(Object resume) {

        return (Resume) resume;
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
