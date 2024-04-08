package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.model.*;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import static ru.javawebinar.basejava.model.ContactType.*;
import static ru.javawebinar.basejava.model.SectionType.*;

public  class ResumeTestData{
    //  Storage storage;
   /* ResumeTestData rtd1 = new ResumeTestData(storage);
     ResumeTestData(Storage storage) {
       this.storage=storage;
         ResumeTestData.create(storage.get())

    }
*/
    public static Resume create( String uuid, String fullName) {
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

        return r;

    }
//rtd1.create(storage.get(uuid),storage.get(fullName));
}
