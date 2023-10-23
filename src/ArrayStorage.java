import javax.swing.*;
import java.util.Arrays;

/**
 * Array based storage for Resumes
 */

public class ArrayStorage {
    Resume[] storage = new Resume[10000];
    int size;

    void clear() {
        Arrays.fill(storage, null);
        size = 0;
    }

    void save(Resume r) {
        size = size + 1;
        storage[size] = r;

    }

    Resume get(String uuid) {
        for (int i = 0; i < size; i++) {
            if (storage[i].uuid == uuid) {
                return storage[i];
            }
        }
    }

    void delete(String uuid) {
        int n = 0;

        for (int i = 0; i < size; i++) {
            if (storage[i].uuid == uuid) {
                n = i;

            }
            if (n != 0) {
                System.arraycopy(storage, n + 1, storage, n, storage.length - 1);
                storage[size] = null;
                size = size - 1;

            }
        }

    }

    /**
     * @return array, contains only Resumes in storage (without null)
     */
    Resume[] getAll() {
        return new Resume[0];
    }

    int size() {
        return size;
    }
}
