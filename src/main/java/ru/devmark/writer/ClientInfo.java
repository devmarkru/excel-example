package ru.devmark.writer;

import java.time.LocalDate;

public class ClientInfo {

    private final String name;
    private final LocalDate birthDate;

    public ClientInfo(String name, LocalDate birthDate) {
        this.name = name;
        this.birthDate = birthDate;
    }

    public String getName() {
        return name;
    }

    public LocalDate getBirthDate() {
        return birthDate;
    }

    @Override
    public String toString() {
        return "Item{" +
                "name='" + name + '\'' +
                ", birthDate=" + birthDate +
                '}';
    }
}
