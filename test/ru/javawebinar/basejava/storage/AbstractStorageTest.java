package ru.javawebinar.basejava.storage;


import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import ru.javawebinar.basejava.exception.ExistStorageException;
import ru.javawebinar.basejava.exception.NotExistStorageException;
import ru.javawebinar.basejava.exception.StorageException;
import ru.javawebinar.basejava.model.Resume;

import java.util.Arrays;
import java.util.List;

public abstract class AbstractStorageTest {
    protected Storage storage;

    AbstractStorageTest(Storage storage) {
        this.storage = storage;
    }

    public static final String UUID_1 = "uuid1";
    public static final String UUID_2 = "uuid2";
    public static final String UUID_3 = "uuid3";

    public static final String fullName_1 = "dduuid1";
    public static final String fullName_2 = "ffuuid2";
    public static final String fullName_3 = "ssuuid3";
    public static final Resume RESUME_1 = new Resume(UUID_1, fullName_1);
    public static final Resume RESUME_2 = new Resume(UUID_2, fullName_2);
    public static final Resume RESUME_3 = new Resume(UUID_3, fullName_3);

    @BeforeEach
   /* public  void create(String uuid, String fullName) {
        Resume r = new Resume(uuid, fullName);

        r.setContacts(PHONE, "+74522");
        r.setContacts(MOBILE, "+56777");
        r.setContacts(HOME_PHONE, "+34343456777");
        r.setContacts(SKYPE, "skipe");
        r.setContacts(MAIL, "+56777");
        r.setContacts(LINKEDID, "Link14444");
        r.setContacts(GITHUB, "2323git");
        r.setContacts(SAVEOVERFLOW, "2323SOF");
        r.setContacts(HOME_PAGE, "wwww.dfdfd");

        r.setSections(PERSONAL, new TextSectoin("не жадный"));
        r.setSections(OBJECTIVE, new TextSectoin("инженер"));
        List<String> dostig = new ArrayList<>();
        dostig.add("Организация команды");
        dostig.add("разработка проектов");
        dostig.add("Налаживание процесса разработки и непрерывной интеграции ERP");
        r.setSections(ACHIEVEMENT, new ListSection(dostig));
        List<String> qvalif = new ArrayList<>();
        qvalif.add("Version control");
        qvalif.add("PostgreSQL");
        qvalif.add("JavaScript");
        r.setSections(QUALIFICATIONS, new ListSection(qvalif));
        Organization org = new Organization("OOO МИТ", "www.OOO MIT", LocalDate.of(2006, 10, 2), LocalDate.of(2011, 5, 10), "Старший разработчик", "Проектирование и разработка онлайн платформы управления проектами Wrike");
        List<Organization> job = new ArrayList<>();
        job.add(org);
        r.setSections(EXPERIENCE, new OrganizationSection(job));
        Organization educat = new Organization("OOO МИТ", "www.OOO MIT", LocalDate.of(2003, 10, 2), LocalDate.of(2005, 5, 10), "Старший разработчик", "Проектирование и разработка онлайн платформы управления проектами Wrike");
        List<Organization> educList = new ArrayList<>();
        educList.add(educat);
        r.setSections(EDUCATION, new OrganizationSection(educList));
        storage.save(r);
        storage.save(r);
        storage.save(r);

    }
   */
    public void setUp() {

        storage.clear();

        storage.save(RESUME_1);
        storage.save(RESUME_2);
        storage.save(RESUME_3);
        ResumeTestData.create(storage.get(UUID_1).getUuid(), storage.get(UUID_1).getFullName());
        ResumeTestData.create(storage.get(UUID_2).getUuid(), storage.get(UUID_2).getFullName());
        ResumeTestData.create(storage.get(UUID_3).getUuid(), storage.get(UUID_3).getFullName());
    }

    @org.junit.jupiter.api.Test
    void size() {
        Assertions.assertEquals(3, storage.size());
    }

    @org.junit.jupiter.api.Test
    void clear() {
        storage.clear();
        Assertions.assertEquals(0, storage.size());
        Assertions.assertThrows(NotExistStorageException.class, () -> storage.get(UUID_1));
    }

    @org.junit.jupiter.api.Test
    void update() {
        Resume newResume = new Resume(UUID_1, "New_Name");
        storage.update(newResume);
        Assertions.assertEquals(newResume, storage.get(UUID_1));
        Assertions.assertThrows(NotExistStorageException.class, () -> storage.get("dummy"));
        //storage.update(new Resume("uuid1","New_Name"));

    }

    @org.junit.jupiter.api.Test
    void getAllSorted() {
        List<Resume> list = storage.getAllSorted();
        Assertions.assertEquals(3, list.size());
        Assertions.assertEquals(list, Arrays.asList(RESUME_1, RESUME_2, RESUME_3));
        //Assertions.assertEquals(new Resume(UUID_1), array.g);
        storage = null;
        Assertions.assertThrows(NullPointerException.class, () -> storage.getAllSorted());
    }

    @org.junit.jupiter.api.Test
    void save() {

        storage.save(new Resume("uuid4", "NameTest"));

        Assertions.assertEquals(4, storage.size());

        Assertions.assertEquals(new Resume("uuid4", "NameTest"), storage.get("uuid4"));

        Assertions.assertThrows(ExistStorageException.class, () -> storage.save(RESUME_1));
        try {
            for (int i = 5; i <= AbstractArrayStorage.STORAGE_LIMIT; i++) {
                storage.save(new Resume("uuid" + i));
            }
        } catch (StorageException e) {
            Assertions.fail();
        }
        storage.clear();

        storage.save(RESUME_1);
        storage.save(RESUME_2);
        storage.save(RESUME_3);

        Assertions.assertThrows(StorageException.class, () -> storage.save(RESUME_1));
    }

    @org.junit.jupiter.api.Test
    void delete() {
        storage.delete("uuid3");
        Assertions.assertEquals(2, storage.size());
        Assertions.assertThrows(NotExistStorageException.class, () -> storage.delete("uuid9"));

    }

    @org.junit.jupiter.api.Test
    void get() {
        Assertions.assertThrows(NotExistStorageException.class, () -> storage.get("uuid9"));
        Assertions.assertEquals(new Resume("uuid1", fullName_1), storage.get("uuid1"));
    }

    @org.junit.jupiter.api.Test
    void getNotExist() {
        Assertions.assertThrows(NotExistStorageException.class, () -> storage.get("dummy"));

    }


}