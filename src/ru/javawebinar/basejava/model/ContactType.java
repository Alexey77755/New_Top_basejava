package ru.javawebinar.basejava.model;

public enum ContactType {
    PHONE("���."),
    MOBILE("���������"),
    HOME_PHONE("�������� �������"),
    SKYPE("Skype"),
    MAIL("�����"),
    LINKEDID("������� LinkedID"),
    GITHUB("������� GitHub"),
    STATCKOVERFLOW("������� Sackoverflow"),
    HOME_PAGE("�������� ��������");
    private final String title;

    ContactType(String title) {
        this.title = title;
    }

    public String getTitle() {
        return title;
    }
}
