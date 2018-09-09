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
