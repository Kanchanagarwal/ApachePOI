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

The Persons sheet has special lock mark and it means that the sheet is protected. Let’s try to modify some cell. You can see the popup message that inform us the sheet is protected and we have to input the password if we want to modify data. For doing this we need to click the unprotect sheet in a context menu, input the password and we can modify data.
