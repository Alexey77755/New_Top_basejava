package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.Resume;

import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MapStorage extends AbstractStorage {
    Map<String, Resume> hashmap = new HashMap<String, Resume>();

    @Override
    protected Object getSearchKey(String uuid) {
        if (hashmap.containsKey(uuid)) {
            return uuid;
        }
        return null;
    }

    @Override
    protected boolean isExist(Object searchKey) {
        return searchKey != null;
    }

    @Override
    protected void doSave(Resume r, Object searchKey) {
        hashmap.put(r.getUuid(), r);
    }

    @Override
    protected void doUpdate(Resume r, Object searchKey) {
        hashmap.put((String) searchKey, r);
    }


    @Override
    protected void doDelete(Object searchKey) {
        hashmap.keySet().removeIf(key -> key.equals((String) searchKey));
    }

    @Override
    protected Resume doGet(Object searchKey) {
        return hashmap.get((String) searchKey);
    }

    @Override
    public void clear() {
        hashmap.clear();
    }

    @Override
    public List<Resume> getAllSorted(Comparator<Resume> r) {
        return hashmap.values().stream().toList();
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
