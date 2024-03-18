package ru.javawebinar.basejava.storage;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import ru.javawebinar.basejava.exception.ExistStorageException;
import ru.javawebinar.basejava.exception.NotExistStorageException;
import ru.javawebinar.basejava.exception.StorageException;
import ru.javawebinar.basejava.model.Resume;

import java.util.Comparator;
import java.util.List;

public abstract class AbstractStorageTest {
    private Storage storage;

    AbstractStorageTest(Storage storage) {
        this.storage = storage;
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
        Assertions.assertEquals(0, storage.size());
        Assertions.assertThrows(NotExistStorageException.class, () -> {
            storage.get(UUID_1);
        });
    }

    @org.junit.jupiter.api.Test
    void update() throws CloneNotSupportedException {
        Assertions.assertThrows(NotExistStorageException.class, () -> {
            storage.update(new Resume("uuid_5"));
        });
        storage.update(new Resume("uuid1"));
        Assertions.assertEquals("uuid1", storage.get(UUID_1).getUuid());
    }

    @org.junit.jupiter.api.Test
    void getAllSorted(Comparator comparatorResume) throws CloneNotSupportedException {
        List<Resume> array = storage.getAllSorted( comparatorResume);
        Assertions.assertEquals(3, storage.getAllSorted( comparatorResume).size());
        //Assertions.assertEquals(new Resume(UUID_1), array.g);
        storage = null;
        Assertions.assertThrows(NullPointerException.class, () -> {
            storage.getAllSorted( comparatorResume);
        });
    }

    @org.junit.jupiter.api.Test
    void save() {

        storage.save(new Resume("uuid4"));

        Assertions.assertEquals(4, storage.size());

        Assertions.assertEquals(new Resume("uuid4"), storage.get("uuid4"));

        Assertions.assertThrows(ExistStorageException.class, () -> {
            storage.save(new Resume("uuid1"));
        });
        try {
            for (int i = 5; i <= AbstractArrayStorage.STORAGE_LIMIT; i++) {
                storage.save(new Resume("uuid" + i));
            }
        } catch (StorageException e) {
            Assertions.fail();
        }
        storage.clear();

        storage.save(new Resume(UUID_1));
        storage.save(new Resume(UUID_2));
        storage.save(new Resume(UUID_3));

        Assertions.assertThrows(StorageException.class, () -> {
            storage.save(new Resume("uuid1"));
        });
    }

    @org.junit.jupiter.api.Test
    void delete() throws CloneNotSupportedException {
        storage.delete("uuid3");
        Assertions.assertEquals(2, storage.size());
        Assertions.assertThrows(NotExistStorageException.class, () -> {
            storage.delete("uuid9");
        });

    }

    @org.junit.jupiter.api.Test
    void get() throws CloneNotSupportedException {
        Assertions.assertThrows(NotExistStorageException.class, () -> {
            storage.get("uuid9");
        });
        Assertions.assertEquals(new Resume("uuid1"), storage.get("uuid1"));
    }

    @org.junit.jupiter.api.Test
    void getNotExist() {
        Assertions.assertThrows(NotExistStorageException.class, () -> {
            storage.get("dummy");
        });

    }
    @org.junit.jupiter.api.Test
    void saveOverflow  () {
        if (storage instanceof ArrayStorage) {
            Assertions.assertThrows(Exception.class, () -> {
                try {
                    for (int i = 4; i <= AbstractArrayStorage.STORAGE_LIMIT; i++) {
                        storage.save(new Resume());
                    }
                } catch (StorageException e) {
                    Assertions.fail();
                }
                storage.save(new Resume());
            });
        }
    }
}