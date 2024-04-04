package ru.javawebinar.basejava.model;

public enum ContactType {
    PHONE("Тел."),
    MOBILE("Мобильный"),
    HOME_PHONE("Домашний телефон"),
    SKYPE("Skype"),
    MAIL("Почта"),
    LINKEDID("Профиль LinkedID"),
    GITHUB("Профиль GitHub"),
    STATCKOVERFLOW("Профиль Sackoverflow"),
    HOME_PAGE("Домашняя страница");
    private final String title;

    ContactType(String title) {
        this.title = title;
    }

    public String getTitle() {
        return title;
    }
}
