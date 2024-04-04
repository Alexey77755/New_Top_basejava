package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.exception.ExistStorageException;
import ru.javawebinar.basejava.exception.NotExistStorageException;
import ru.javawebinar.basejava.model.Resume;

import java.util.Collections;
import java.util.List;
import java.util.logging.Logger;

public abstract class AbstractStorage<SK> implements Storage {
private static final Logger LOG = Logger.getLogger(AbstractStorage.class.getName());
    //public static Comparator<Resume> comparatorResume = Comparator.comparing(Resume::getFullName).thenComparing(Resume::getUuid);

    protected abstract SK getSearchKey(String uuid);

    protected abstract void doSave(Resume r, SK searchKey);

    protected abstract void doUpdate(Resume r, SK searchKey);

    protected abstract boolean isExist(SK searchKey);

    protected abstract void doDelete(SK searchKey);


    protected abstract Resume doGet(SK searchKey);
    protected abstract List<Resume> doCopyAll();

    @Override
    public void update(Resume r) {
        LOG.info("update "+ r);
        SK searchKey = getExistedSearchKey(r.getUuid());
        doUpdate(r, searchKey);

    }

    public void save(Resume r) {
        LOG.info("save "+ r);
        SK searchKey = getNotExistedSearchKey(r.getUuid());
        doSave(r, searchKey);
    }


    @Override
    public void delete(String uuid) {
        LOG.info("delete "+ uuid);
        SK searchKey = getExistedSearchKey(uuid);
        doDelete(searchKey);
    }

    @Override
    public Resume get(String uuid) {
        LOG.info("get "+ uuid);
        SK searchKey = getExistedSearchKey(uuid);
        return doGet(searchKey);
    }


    private SK getExistedSearchKey(String uuid) {
        SK searchKey = getSearchKey(uuid);
        if (!isExist(searchKey)) {
            LOG.warning("Resume " +uuid + " not exist");
            throw new NotExistStorageException(uuid);
        }
        return searchKey;
    }

    private SK getNotExistedSearchKey(String uuid) {
        SK searchKey = getSearchKey(uuid);
        if (isExist(searchKey)) {
            LOG.warning("Resume " +uuid + " exist");
            throw new ExistStorageException(uuid);
        }
        return searchKey;
    }

    @Override
    public List<Resume> getAllSorted() {
        LOG.info("getAllSorted");
        List<Resume> list =doCopyAll();
        Collections.sort(list);
        return list;
    }




}
