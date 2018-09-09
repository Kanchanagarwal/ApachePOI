# 10 - Protect sheet by password

The apache poi provides the protect sheet feature. The goal of the feature is to reject users from accidentally or deliberately changing, moving, or deleting data in a worksheet, you can protect it with a password.

## Dependencies

```
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>3.17</version>
</dependency>
```
## Result

The "Persons" sheet has a special lock mark and it means that the sheet is protected. When a user tries to modify the sheet the popup message is shown that informs a user that the sheet is protected and we have to input the password to modify data. 

<img width="1280" alt="screen shot 2018-09-09 at 11 19 16" src="https://user-images.githubusercontent.com/5372875/45263072-299fe580-b423-11e8-8151-3a994cab86b6.png">
