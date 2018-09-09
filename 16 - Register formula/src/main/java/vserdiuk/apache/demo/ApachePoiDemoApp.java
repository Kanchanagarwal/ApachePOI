/*
    Licensed to the Apache Software Foundation (ASF) under one
    or more contributor license agreements.  See the NOTICE file
    distributed with this work for additional information
    regarding copyright ownership.  The ASF licenses this file
    to you under the Apache License, Version 2.0 (the
    "License"); you may not use this file except in compliance
    with the License.  You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing,
    software distributed under the License is distributed on an
    "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
    KIND, either express or implied.  See the License for the
    specific language governing permissions and limitations
    under the License.
*/

package vserdiuk.apache.demo;

import vserdiuk.apache.demo.excel.PersonExcelWriter;
import vserdiuk.apache.demo.model.Person;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class ApachePoiDemoApp {
    public static void main(String[] args) {
        List<Person> personList = new ArrayList<>();

        Person person1 = new Person();
        person1.setFirstName("John");
        person1.setLastName("Smith");
        person1.setBirthday(LocalDate.parse("1990-05-10"));
        person1.setEmail("john.smith@email.com");
        person1.setPhoneNumber("123456789");
        person1 .setMarried(true);

        Person person2 = new Person();
        person2.setFirstName("Mary");
        person2.setLastName("Brown");
        person2.setBirthday(LocalDate.parse("2007-05-10"));
        person2.setEmail("mary.brown@email.com");
        person2.setPhoneNumber("987654321");
        person2.setMarried(false);

        Person person3 = new Person();
        person3.setFirstName("Mason");
        person3.setLastName("Johnson");
        person3.setBirthday(LocalDate.parse("2000-05-10"));
        person3.setEmail("mason.johnson@email.com");
        person3.setPhoneNumber("987654321");
        person3.setMarried(false);

        Person person4 = new Person();
        person4.setFirstName("Matthew");
        person4.setLastName("Williams");
        person4.setBirthday(LocalDate.parse("2001-05-10"));
        person4.setEmail("matthew.williams@email.com");
        person4.setPhoneNumber("987654321");
        person4.setMarried(false);

        Person person5 = new Person();
        person5.setFirstName("James");
        person5.setLastName("Jones");
        person5.setBirthday(LocalDate.parse("1990-05-10"));
        person5.setEmail("james.jones@email.com");
        person5.setPhoneNumber("987654321");
        person5.setMarried(false);

        Person person6 = new Person();
        person6.setFirstName("Mary");
        person6.setLastName("Garcia");
        person6.setBirthday(LocalDate.parse("2003-05-10"));
        person6.setEmail("mary.garcia@email.com");
        person6.setPhoneNumber("987654321");
        person6.setMarried(false);

        Person person7 = new Person();
        person7.setFirstName("Mary");
        person7.setLastName("Davis");
        person7.setBirthday(LocalDate.parse("1980-05-10"));
        person7.setEmail("mary.davis@email.com");
        person7.setPhoneNumber("987654321");
        person7.setMarried(false);

        Person person8 = new Person();
        person8.setFirstName("Mia");
        person8.setLastName("Martinez");
        person8.setBirthday(LocalDate.parse("2000-05-10"));
        person8.setEmail("mia.martinez@email.com");
        person8.setPhoneNumber("987654321");
        person8.setMarried(false);

        Person person9 = new Person();
        person9.setFirstName("Victoria");
        person9.setLastName("Lopez");
        person9.setBirthday(LocalDate.parse("2000-05-10"));
        person9.setEmail("victoria.lopez@email.com");
        person9.setPhoneNumber("987654321");
        person9.setMarried(false);

        Person person10 = new Person();
        person10.setFirstName("Alexis");
        person10.setLastName("Anderson");
        person10.setBirthday(LocalDate.parse("1972-05-10"));
        person10.setEmail("Alexis.Anderson@email.com");
        person10.setPhoneNumber("987654321");
        person10.setMarried(false);

        personList.add(person1);
        personList.add(person2);
        personList.add(person3);
        personList.add(person4);
        personList.add(person5);
        personList.add(person6);
        personList.add(person7);
        personList.add(person8);
        personList.add(person9);
        personList.add(person10);

        PersonExcelWriter writer = new PersonExcelWriter();
        writer.write("demo.xlsx", personList);
    }
}
