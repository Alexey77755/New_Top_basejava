package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MapResumeStorage extends AbstractStorage<Resume> {
    Map<String, Resume> hashmap = new HashMap<String, Resume>();

    @Override
    protected Resume getSearchKey(String uuid) {

        return hashmap.get(uuid);
        }


    @Override
    protected boolean isExist(Resume resume) {
        return resume!=null;
    }



    @Override
    protected void doSave(Resume r, Resume resume) {

        hashmap.put(r.getUuid(), r);
    }

    @Override
    protected void doUpdate(Resume r, Resume resume) {

        hashmap.put(r.getUuid(), r);
    }


    @Override
    protected void doDelete(Resume resume) {
        hashmap.keySet().remove(resume.getUuid());
    }

    @Override
    protected Resume doGet(Resume resume) {

        return resume;
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
