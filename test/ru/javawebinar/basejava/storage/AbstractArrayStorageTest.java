package ru.javawebinar.basejava.storage;

import org.junit.jupiter.api.Assertions;
import ru.javawebinar.basejava.exception.StorageException;
import ru.javawebinar.basejava.model.Resume;

//@ExtendWith(TestRailExtension.class)
public abstract class AbstractArrayStorageTest extends AbstractStorageTest {
    AbstractArrayStorageTest(Storage storage) {
        super(storage);
     }

    @org.junit.jupiter.api.Test
    void saveOverflow() {

        Assertions.assertThrows(Exception.class, () -> {
            try {
                for (int i = 4; i <= AbstractArrayStorage.STORAGE_LIMIT; i++) {
                    storage.save(new Resume("name"));
                }
            } catch (StorageException e) {
                Assertions.fail();
            }
            storage.save(new Resume("OverFlow"));
        });

    }
}