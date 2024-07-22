package org.example.domain;

public class Person {
    private String firstName;
    private String lastName;
    private String rut;

    public Person(String firstName, String lastName, String rut) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.rut = rut;
    }

    public String getFirstName() {
        return firstName;
    }

    public String getLastName() {
        return lastName;
    }

    public String getRut() {
        return rut;
    }
}

