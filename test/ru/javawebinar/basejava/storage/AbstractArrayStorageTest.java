package ru.javawebinar.basejava.storage;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import ru.javawebinar.basejava.exception.ExistStorageException;
import ru.javawebinar.basejava.exception.NotExistStorageException;
import ru.javawebinar.basejava.model.Resume;

public abstract class  AbstractArrayStorageTest {
    private Storage storage ;
    AbstractArrayStorageTest(Storage storage){
        this.storage=storage;
    }

    public static final String UUID_1 = "uuid1";
    public static final String UUID_2 = "uuid2";
    public static final String UUID_3 = "uuid3";

    @BeforeEach
    public void setUp() throws Exception {
        storage.clear();
        storage.save(new Resume(UUID_1));
        storage.save(new Resume(UUID_2));
        storage.save(new Resume(UUID_3));
    }

    @org.junit.jupiter.api.Test
    void size() {
        Assertions.assertEquals(3, storage.size());
    }

    @org.junit.jupiter.api.Test
    void clear() {
        storage.clear();
        Assertions.assertEquals(0, storage.getAll().length);
        Assertions.assertEquals(0, storage.size());
        Assertions.assertThrows(NotExistStorageException.class,() -> { storage.get(UUID_1);});
    }

    @org.junit.jupiter.api.Test
    void update() {
        Assertions.assertThrows(NotExistStorageException.class,() -> { storage.update(new Resume("uuid_5"));});
        storage.update(new Resume("uuid1"));
        Assertions.assertEquals("uuid1",storage.get(UUID_1).getUuid());
    }

    @org.junit.jupiter.api.Test
    void getAll() {
        Assertions.assertEquals(3, storage.getAll().length);
        storage=null;
        Assertions.assertThrows(NullPointerException.class,() -> { storage.getAll();});
    }

    @org.junit.jupiter.api.Test
    void save() {

        Assertions.assertThrows(ExistStorageException.class,() -> { storage.save(new Resume("uuid1"));});

    }

    @org.junit.jupiter.api.Test
    void delete() {
    }

    @org.junit.jupiter.api.Test
    void get() {
    }
    @org.junit.jupiter.api.Test
    void getNotExist() {
        Assertions.assertThrows(NotExistStorageException.class,() -> { storage.get("dummy");});

    }
}