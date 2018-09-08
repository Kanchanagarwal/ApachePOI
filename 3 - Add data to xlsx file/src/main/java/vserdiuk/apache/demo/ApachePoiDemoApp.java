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

        personList.add(person1);
        personList.add(person2);

        PersonExcelWriter writer = new PersonExcelWriter();
        writer.write("demo.xlsx", personList);
    }
}
