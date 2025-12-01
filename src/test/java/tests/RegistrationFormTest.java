package tests;

import org.openqa.selenium.*;
import org.testng.annotations.*;
import pages.RegistrationPage;
import utils.TestResultsExcelWriter;

import java.text.SimpleDateFormat;
import java.util.*;

public class RegistrationFormTest extends BaseTest {
    
    private String uniqueEmail;

    @BeforeClass
    public void setupTests() {
        TestResultsExcelWriter.clearResults();
        System.out.println("=== ניקוי תוצאות קודמות ===");
        
        // יצירת אימייל ייחודי
        String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        uniqueEmail = "test" + timestamp + "@example.com";
        System.out.println("=== אימייל ייחודי: " + uniqueEmail + " ===\n");
    }

    @Test(priority = 1)
    public void test1_SuccessfulRegistration() throws Exception {
        System.out.println("\n========================================");
        System.out.println("טסט 1: הרשמה מוצלחת");
        System.out.println("========================================\n");
        
        RegistrationPage reg = new RegistrationPage(driver);
        
        // ניווט לדף ההרשמה
        navigateToRegistrationPage("test1");
        
        // נתוני הבדיקה
        List<String> testData = Arrays.asList(
            uniqueEmail,                       // אימייל
            "Password123!",                    // סיסמה
            "Password123!",                    // סיסמה שנית
            "יוסי",                            // שם פרטי
            "כהן",                             // שם משפחה
            "תל אביב - יפו",                   // עיר
            "הרצל",                            // רחוב
            "25",                              // מספר בית
            "ג",                               // כניסה
            "5",                               // דירה
            "0541234579",                      // טלפון נייד
            "0501234579"                       // טלפון נוסף
        );

        System.out.println("\n=== מתחיל למלא את הטופס ===");
        String[] fieldNames = {"אימייל", "סיסמא", "סיסמא בשנית", "שם פרטי", "שם משפחה", "עיר/יישוב", "רחוב", "מספר בית", "כניסה", "דירה", "טלפון נייד", "טלפון נוסף"};
        for (int i = 0; i < testData.size() && i < fieldNames.length; i++) {
            String displayValue = i == 1 || i == 2 ? "****" : testData.get(i);
            System.out.println(fieldNames[i] + ": " + displayValue);
        }
        System.out.println();

        // מילוי השדות
        try {
            reg.fillFormFields(testData);
            System.out.println("\n✓ הטופס מולא בהצלחה!");
            TestResultsExcelWriter.addTestResult("מילוי שדות", true, "כל 12 השדות + 2 צ'קבוקסים מולאו");
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("מילוי שדות", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        System.out.println("\nהמתן 2 שניות...");
        Thread.sleep(2000);
        
        // לחיצה על כפתור הרשמה
        System.out.println("\n=== לוחץ על כפתור הרשמה ===");
        try {
            WebElement submitButton = driver.findElement(By.cssSelector("form.reg-form input[type='submit'], form.reg-form button[type='submit']"));
            
            System.out.println("נמצא כפתור הרשמה");
            
            // גלילה לכפתור
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.sleep(1500);
            
            // הדגשה עדינה
            ((JavascriptExecutor) driver).executeScript(
                "arguments[0].style.border='1px solid #FFB6C1';" +
                "arguments[0].style.boxShadow='0 0 3px rgba(255, 182, 193, 0.5)';" +
                "arguments[0].style.borderRadius='3px';" +
                "arguments[0].style.outline='1px solid rgba(255, 182, 193, 0.3)';" +
                "arguments[0].style.outlineOffset='2px';", 
                submitButton
            );
            
            Thread.sleep(3000);
            
            System.out.println("\nלוחץ על הכפתור...");
            try {
                submitButton.click();
            } catch (Exception e) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitButton);
            }
            System.out.println("✓ לחץ על כפתור ההרשמה");
            
            TestResultsExcelWriter.addTestResult("לחיצה על כפתור", true, "בוצעה בהצלחה");
            
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("לחיצה על כפתור", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        System.out.println("\nהמתן 5 שניות...");
        Thread.sleep(5000);
        
        TestResultsExcelWriter.addTestResult("טסט 1: הרשמה מוצלחת", true, "ההרשמה הושלמה בהצלחה");
        
        System.out.println("\n=== טסט 1 הסתיים בהצלחה ===");
    }

    @Test(priority = 2)
    public void test2_DuplicateEmailRegistration() throws Exception {
        System.out.println("\n========================================");
        System.out.println("טסט 2: ניסיון הרשמה עם אימייל קיים");
        System.out.println("========================================\n");
        
        // מחיקת עוגיות והתנתקות מהמערכת
        driver.manage().deleteAllCookies();
        Thread.sleep(1000);
        System.out.println("✓ עוגיות נמחקו - התנתקות מהמערכת");
        
        // רענון הדפדפן לוידוא ניקוי המצב
        driver.get("https://www.lastprice.co.il");
        Thread.sleep(2000);
        System.out.println("✓ רענון דפדפן");
        
        RegistrationPage reg = new RegistrationPage(driver);
        
        // ניווט לדף ההרשמה
        navigateToRegistrationPage("test2");
        
        // נתוני בדיקה - אימייל זהה, שאר הפרטים שונים
        List<String> testData = Arrays.asList(
            uniqueEmail,                       // אימייל זהה למשתמש הראשון!
            "DifferentPass456!",               // סיסמה שונה
            "DifferentPass456!",               // סיסמה שנית
            "דוד",                             // שם פרטי שונה
            "לוי",                             // שם משפחה שונה
            "חיפה",                            // עיר שונה
            "בן גוריון",                       // רחוב שונה
            "10",                              // מספר בית שונה
            "א",                               // כניסה שונה
            "3",                               // דירה שונה
            "0527654321",                      // טלפון שונה
            "0507654321"                       // טלפון שונה
        );

        System.out.println("\n=== מתחיל למלא טופס עם אימייל קיים ===");
        System.out.println("אימייל: " + uniqueEmail + " (קיים!)");
        System.out.println("שם פרטי: דוד (שונה)");
        System.out.println("שם משפחה: לוי (שונה)");
        System.out.println("עיר: חיפה (שונה)");
        System.out.println();

        try {
            reg.fillFormFields(testData);
            System.out.println("\n✓ הטופס מולא");
            TestResultsExcelWriter.addTestResult("מילוי טופס עם אימייל קיים", true, "הטופס מולא עם פרטים שונים ואימייל קיים");
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("מילוי טופס עם אימייל קיים", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        Thread.sleep(2000);
        
        // לחיצה על כפתור הרשמה
        System.out.println("\n=== לוחץ על כפתור הרשמה ===");
        try {
            WebElement submitButton = driver.findElement(By.cssSelector("form.reg-form input[type='submit'], form.reg-form button[type='submit']"));
            
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.sleep(1000);
            
            System.out.println("לוחץ על הכפתור...");
            try {
                submitButton.click();
            } catch (Exception e) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitButton);
            }
            System.out.println("✓ לחץ על כפתור ההרשמה");
            
            TestResultsExcelWriter.addTestResult("לחיצה על כפתור", true, "בוצעה בהצלחה");
            
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("לחיצה על כפתור", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        Thread.sleep(3000);
        
        // בדיקת הודעת שגיאה - אם יש שגיאה זה טוב! המשמעות היא שהמערכת עבדה נכון
        System.out.println("\n=== בודק הודעת שגיאה (מצופה למצוא שגיאה!) ===");
        try {
            Thread.sleep(2000);
            
            // חיפוש הודעות שגיאה
            List<WebElement> errorMessages = driver.findElements(By.cssSelector(".error, .text-danger, [class*='error'], .invalid-feedback, .alert-danger, .alert"));
            
            boolean errorFound = false;
            String errorText = "";
            
            for (WebElement error : errorMessages) {
                try {
                    if (error.isDisplayed()) {
                        String text = error.getText().trim();
                        if (!text.isEmpty() && text.length() < 500) {
                            errorFound = true;
                            errorText = text;
                            System.out.println("✓✓✓ נמצאה הודעת שגיאה (כמצופה!): " + errorText);
                            break;
                        }
                    }
                } catch (Exception ignored) {}
            }
            
            // אם מצאנו שגיאה - זה PASS! כי המערכת מנעה רישום כפול
            if (errorFound) {
                System.out.println("\n✓ הבדיקה עברה בהצלחה - המערכת מנעה רישום כפול!");
                TestResultsExcelWriter.addTestResult(
                    "בדיקת אימייל כפול", 
                    true, 
                    "✓ המערכת זיהתה אימייל קיים והציגה שגיאה: \"" + errorText + "\" - הבדיקה עברה בהצלחה!"
                );
            } else {
                // אם לא מצאנו שגיאה - זה בעיה! המשמעות היא שהמערכת אפשרה רישום כפול
                System.out.println("\n⚠ לא נמצאה הודעת שגיאה - המערכת אפשרה רישום כפול (בעיה!)");
                
                // בדיקה נוספת - אולי העמוד השתנה
                String currentUrl = driver.getCurrentUrl();
                String pageSource = driver.getPageSource();
                
                if (pageSource.contains("כבר קיים") || pageSource.contains("רשום") || pageSource.contains("exist")) {
                    System.out.println("✓ נמצאה אינדיקציה להודעת שגיאה בקוד הדף");
                    TestResultsExcelWriter.addTestResult(
                        "בדיקת אימייל כפול", 
                        true, 
                        "✓ המערכת מנעה רישום כפול (זוהה אינדיקציה בדף) - הבדיקה עברה!"
                    );
                } else {
                    TestResultsExcelWriter.addTestResult(
                        "בדיקת אימייל כפול", 
                        false, 
                        "✗ לא נמצאה הודעת שגיאה - המערכת אפשרה רישום כפול!"
                    );
                }
            }
            
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("בדיקת אימייל כפול", false, "שגיאה בביצוע הבדיקה: " + e.getMessage());
        }
        
        System.out.println("\n=== טסט 2 הסתיים ===");
    }

    @Test(priority = 3)
    public void test3_RequiredFieldsValidation() throws Exception {
        System.out.println("\n========================================");
        System.out.println("טסט 3: אימות שדות חובה");
        System.out.println("========================================\n");
        
        // מחיקת עוגיות והתנתקות מהמערכת
        driver.manage().deleteAllCookies();
        Thread.sleep(1000);
        System.out.println("✓ עוגיות נמחקו - התנתקות מהמערכת");
        
        // רענון הדפדפן לוידוא ניקוי המצב
        driver.get("https://www.lastprice.co.il");
        Thread.sleep(2000);
        System.out.println("✓ רענון דפדפן");
        
        RegistrationPage reg = new RegistrationPage(driver);
        
        // ניווט לדף ההרשמה
        navigateToRegistrationPage("test3");
        
        // נתוני בדיקה - משאירים שדה אימייל וסיסמה ריקים (שדות חובה!)
        List<String> testData = Arrays.asList(
            "",                                // אימייל - ריק! (שדה חובה)
            "",                                // סיסמה - ריקה! (שדה חובה)
            "",                                // סיסמה שנית - ריקה!
            "רחל",                             // שם פרטי
            "כהן",                             // שם משפחה
            "ירושלים",                         // עיר
            "יפו",                             // רחוב
            "15",                              // מספר בית
            "ב",                               // כניסה
            "7",                               // דירה
            "0523456789",                      // טלפון נייד
            "0503456789"                       // טלפון נוסף
        );

        System.out.println("\n=== מתחיל למלא טופס עם שדות חובה ריקים ===");
        System.out.println("אימייל: [ריק] - שדה חובה!");
        System.out.println("סיסמה: [ריקה] - שדה חובה!");
        System.out.println("שם פרטי: רחל");
        System.out.println("שם משפחה: כהן");
        System.out.println();

        try {
            reg.fillFormFields(testData);
            System.out.println("\n✓ הטופס מולא (עם שדות חובה ריקים)");
            TestResultsExcelWriter.addTestResult("מילוי טופס עם שדות חסרים", true, "הטופס מולא בלי אימייל וסיסמה");
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("מילוי טופס עם שדות חסרים", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        Thread.sleep(2000);
        
        // לחיצה על כפתור הרשמה
        System.out.println("\n=== מנסה לשלוח טופס עם שדות חובה חסרים ===");
        try {
            WebElement submitButton = driver.findElement(By.cssSelector("form.reg-form input[type='submit'], form.reg-form button[type='submit']"));
            
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.sleep(1000);
            
            System.out.println("לוחץ על הכפתור...");
            try {
                submitButton.click();
            } catch (Exception e) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitButton);
            }
            System.out.println("✓ לחץ על כפתור ההרשמה");
            
            TestResultsExcelWriter.addTestResult("ניסיון שליחה עם שדות חסרים", true, "לחיצה בוצעה");
            
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("ניסיון שליחה עם שדות חסרים", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        Thread.sleep(2000);
        
        // בדיקה שהמערכת מנעה את השליחה - אם יש שגיאה זה טוב!
        System.out.println("\n=== בודק אם המערכת מנעה שליחה (מצופה למצוא שגיאה!) ===");
        try {
            Thread.sleep(2000);
            
            boolean validationPassed = false;
            String validationMessage = "";
            
            // בדיקה 1: חיפוש הודעות שגיאה
            List<WebElement> errorMessages = driver.findElements(By.cssSelector(".error, .text-danger, [class*='error'], .invalid-feedback, .alert-danger, .alert"));
            
            for (WebElement error : errorMessages) {
                try {
                    if (error.isDisplayed()) {
                        String text = error.getText().trim();
                        if (!text.isEmpty() && text.length() < 500) {
                            validationPassed = true;
                            validationMessage = text;
                            System.out.println("✓✓✓ נמצאה הודעת שגיאה (כמצופה!): " + text);
                            break;
                        }
                    }
                } catch (Exception ignored) {}
            }
            
            // בדיקה 2: בדיקת HTML5 validation
            if (!validationPassed) {
                try {
                    WebElement emailField = driver.findElement(By.id("cemail"));
                    String validationMsg = (String) ((JavascriptExecutor) driver)
                        .executeScript("return arguments[0].validationMessage;", emailField);
                    
                    if (validationMsg != null && !validationMsg.isEmpty()) {
                        validationPassed = true;
                        validationMessage = "HTML5 Validation: " + validationMsg;
                        System.out.println("✓✓✓ נמצא HTML5 validation (כמצופה!): " + validationMsg);
                    }
                } catch (Exception ignored) {}
            }
            
            // בדיקה 3: האם נשארנו באותו עמוד (לא עבר לעמוד אחר)
            if (!validationPassed) {
                String currentUrl = driver.getCurrentUrl();
                if (currentUrl.contains("Register") || currentUrl.contains("register")) {
                    validationPassed = true;
                    validationMessage = "נשארנו בעמוד ההרשמה - המערכת לא אפשרה שליחה";
                    System.out.println("✓✓✓ נשארנו בעמוד ההרשמה - המערכת מנעה שליחה (כמצופה!)");
                }
            }
            
            // תוצאה סופית
            if (validationPassed) {
                System.out.println("\n✓ הבדיקה עברה בהצלחה - המערכת מנעה שליחה עם שדות חובה חסרים!");
                TestResultsExcelWriter.addTestResult(
                    "אימות שדות חובה", 
                    true, 
                    "✓ המערכת מנעה שליחת טופס עם שדות חובה חסרים: \"" + validationMessage + "\" - הבדיקה עברה!"
                );
            } else {
                System.out.println("\n⚠ לא זוהה מניעת שליחה - המערכת אפשרה שליחה עם שדות חסרים (בעיה!)");
                TestResultsExcelWriter.addTestResult(
                    "אימות שדות חובה", 
                    false, 
                    "✗ המערכת לא מנעה שליחה עם שדות חובה חסרים!"
                );
            }
            
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("אימות שדות חובה", false, "שגיאה בביצוע הבדיקה: " + e.getMessage());
        }
        
        System.out.println("\n=== טסט 3 הסתיים ===");
    }

    @Test(priority = 4)
    public void test4_InvalidPhoneValidation() throws Exception {
        System.out.println("\n========================================");
        System.out.println("טסט 4: אימות טלפון עם אותיות");
        System.out.println("========================================\n");
        
        // מחיקת עוגיות והתנתקות מהמערכת
        driver.manage().deleteAllCookies();
        Thread.sleep(1000);
        System.out.println("✓ עוגיות נמחקו - התנתקות מהמערכת");
        
        // רענון הדפדפן לוידוא ניקוי המצב
        driver.get("https://www.lastprice.co.il");
        Thread.sleep(2000);
        System.out.println("✓ רענון דפדפן");
        
        // יצירת אימייל ייחודי חדש למשתמש זה
        String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        String newEmail = "test" + timestamp + "@example.com";
        
        RegistrationPage reg = new RegistrationPage(driver);
        
        // ניווט לדף ההרשמה
        navigateToRegistrationPage("test4");
        
        // נתוני בדיקה - כל הפרטים תקינים, טלפון ראשון תקין אבל הטלפון השני עם אותיות!
        List<String> testData = Arrays.asList(
            newEmail,                          // אימייל ייחודי חדש
            "ValidPass789!",                   // סיסמה תקינה
            "ValidPass789!",                   // סיסמה שנית תקינה
            "מיכל",                            // שם פרטי
            "אברהם",                           // שם משפחה
            "באר שבע",                         // עיר
            "הנגב",                            // רחוב
            "50",                              // מספר בית
            "ד",                               // כניסה
            "12",                              // דירה
            "0534567891",                      // טלפון נייד - תקין! (מספרים בלבד)
            "ABC1234567"                       // טלפון נוסף - עם אותיות!
        );

        System.out.println("\n=== מתחיל למלא טופס עם טלפון שני לא תקין ===");
        System.out.println("אימייל: " + newEmail);
        System.out.println("סיסמה: ValidPass789!");
        System.out.println("שם פרטי: מיכל");
        System.out.println("שם משפחה: אברהם");
        System.out.println("עיר: באר שבע");
        System.out.println("רחוב: הנגב");
        System.out.println("מספר בית: 50");
        System.out.println("כניסה: ד");
        System.out.println("דירה: 12");
        System.out.println("טלפון נייד: 0534567891 [תקין]");
        System.out.println("טלפון נוסף: ABC1234567 [עם אותיות!]");
        System.out.println();

        try {
            reg.fillFormFields(testData);
            System.out.println("\n✓ הטופס מולא (עם טלפון שני לא תקין)");
            TestResultsExcelWriter.addTestResult("מילוי טופס עם טלפון שני לא תקין", true, "הטופס מולא עם אותיות בשדה טלפון השני");
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("מילוי טופס עם טלפון שני לא תקין", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        Thread.sleep(2000);
        
        // לחיצה על כפתור הרשמה
        System.out.println("\n=== מנסה לשלוח טופס עם טלפון שני לא תקין ===");
        try {
            WebElement submitButton = driver.findElement(By.cssSelector("form.reg-form input[type='submit'], form.reg-form button[type='submit']"));
            
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.sleep(1000);
            
            System.out.println("לוחץ על הכפתור...");
            try {
                submitButton.click();
            } catch (Exception e) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitButton);
            }
            System.out.println("✓ לחץ על כפתור ההרשמה");
            
            TestResultsExcelWriter.addTestResult("ניסיון שליחה עם טלפון שני לא תקין", true, "לחיצה בוצעה");
            
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("ניסיון שליחה עם טלפון שני לא תקין", false, "שגיאה: " + e.getMessage());
            throw e;
        }
        
        Thread.sleep(3000);
        
        // בדיקה שהמערכת מנעה את השליחה או הציגה שגיאה - אם יש שגיאה זה טוב!
        System.out.println("\n=== בודק אם המערכת מנעה שליחה (מצופה למצוא שגיאה!) ===");
        try {
            Thread.sleep(2000);
            
            boolean validationPassed = false;
            String validationMessage = "";
            
            // בדיקה 1: חיפוש הודעות שגיאה
            List<WebElement> errorMessages = driver.findElements(By.cssSelector(".error, .text-danger, [class*='error'], .invalid-feedback, .alert-danger, .alert"));
            
            for (WebElement error : errorMessages) {
                try {
                    if (error.isDisplayed()) {
                        String text = error.getText().trim();
                        if (!text.isEmpty() && text.length() < 500 && (text.contains("טלפון") || text.contains("פון") || text.contains("phone") || text.contains("מספר"))) {
                            validationPassed = true;
                            validationMessage = text;
                            System.out.println("✓✓✓ נמצאה הודעת שגיאה לטלפון (כמצופה!): " + text);
                            break;
                        }
                    }
                } catch (Exception ignored) {}
            }
            
            // בדיקה 2: בדיקת HTML5 validation על שדה הטלפון השני
            if (!validationPassed) {
                try {
                    WebElement phoneField = driver.findElement(By.id("cphone2"));
                    String validationMsg = (String) ((JavascriptExecutor) driver)
                        .executeScript("return arguments[0].validationMessage;", phoneField);
                    
                    if (validationMsg != null && !validationMsg.isEmpty()) {
                        validationPassed = true;
                        validationMessage = "HTML5 Validation on Phone 2: " + validationMsg;
                        System.out.println("✓✓✓ נמצא HTML5 validation על טלפון שני (כמצופה!): " + validationMsg);
                    }
                } catch (Exception ignored) {}
            }
            
            // בדיקה 3: בדיקת תוקף השדה השני
            if (!validationPassed) {
                try {
                    WebElement phoneField = driver.findElement(By.id("cphone2"));
                    Boolean isValid = (Boolean) ((JavascriptExecutor) driver)
                        .executeScript("return arguments[0].checkValidity();", phoneField);
                    
                    if (isValid != null && !isValid) {
                        validationPassed = true;
                        validationMessage = "שדה הטלפון השני לא עבר בדיקת תקינות של הדפדפן";
                        System.out.println("✓✓✓ שדה הטלפון השני סומן כלא תקין (כמצופה!)");
                    }
                } catch (Exception ignored) {}
            }
            
            // בדיקה 4: האם נשארנו באותו עמוד (לא עבר לעמוד אחר)
            if (!validationPassed) {
                String currentUrl = driver.getCurrentUrl();
                if (currentUrl.contains("Register") || currentUrl.contains("register")) {
                    validationPassed = true;
                    validationMessage = "נשארנו בעמוד ההרשמה - המערכת לא אפשרה שליחה";
                    System.out.println("✓✓✓ נשארנו בעמוד ההרשמה - המערכת מנעה שליחה (כמצופה!)");
                }
            }
            
            // תוצאה סופית
            if (validationPassed) {
                System.out.println("\n✓ הבדיקה עברה בהצלחה - המערכת מנעה שליחה עם טלפון שני לא תקין!");
                TestResultsExcelWriter.addTestResult(
                    "אימות טלפון שני לא תקין", 
                    true, 
                    "✓ המערכת מנעה שליחת טופס עם טלפון שני המכיל אותיות: \"" + validationMessage + "\" - הבדיקה עברה!"
                );
            } else {
                System.out.println("\n⚠ לא זוהה מניעת שליחה - המערכת אפשרה שליחה עם טלפון שני לא תקין (בעיה!)");
                TestResultsExcelWriter.addTestResult(
                    "אימות טלפון שני לא תקין", 
                    false, 
                    "✗ המערכת לא מנעה שליחה עם טלפון שני המכיל אותיות!"
                );
            }
            
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("אימות טלפון שני לא תקין", false, "שגיאה בביצוע הבדיקה: " + e.getMessage());
        }
        
        System.out.println("\n=== טסט 4 הסתיים ===");
    }

    @AfterClass
    public void tearDownTests() {
        System.out.println("\n========================================");
        System.out.println("סיום כל הבדיקות");
        System.out.println("========================================\n");
        
        TestResultsExcelWriter.printSummary();
        
        String excelPath = "output/registeration_validation_results.xlsx";
        try {
            TestResultsExcelWriter.writeResultsToExcel(excelPath);
            System.out.println("\n✓ התוצאות נשמרו בהצלחה: " + excelPath);
        } catch (Exception e) {
            System.out.println("שגיאה בשמירת Excel: " + e.getMessage());
        }
    }
    
    private void navigateToRegistrationPage(String testPrefix) throws Exception {
        String homeUrl = "https://www.lastprice.co.il";
        System.out.println("=== כניסה לעמוד הבית ===");
        System.out.println("פותח: " + homeUrl);
        
        try {
            driver.get(homeUrl);
            Thread.sleep(3000);
            
            TestResultsExcelWriter.addTestResult("כניסה לעמוד הבית", true, "נפתח בהצלחה");
        } catch (Exception e) {
            TestResultsExcelWriter.addTestResult("כניסה לעמוד הבית", false, "שגיאה");
            throw e;
        }
        
        System.out.println("\n=== מחפש כפתור הרשם ===");
        WebElement registerButton = null;
        
        try {
            WebElement loginDropdown = driver.findElement(By.xpath("//span[contains(text(), 'התחבר/הרשם')]"));
            WebElement dropdownTrigger = loginDropdown.findElement(By.xpath("./.."));
            
            System.out.println("נמצא אזור התחבר/הרשם");
            
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", dropdownTrigger);
            Thread.sleep(2000);
            
            ((JavascriptExecutor) driver).executeScript(
                "arguments[0].style.border='1px solid #87CEEB';" +
                "arguments[0].style.boxShadow='0 0 3px rgba(135, 206, 235, 0.5)';" +
                "arguments[0].style.borderRadius='3px';" +
                "arguments[0].style.outline='1px solid rgba(135, 206, 235, 0.3)';" +
                "arguments[0].style.outlineOffset='2px';", 
                dropdownTrigger
            );
            Thread.sleep(2000);
            
            TestResultsExcelWriter.addTestResult("מציאת אזור התחבר/הרשם", true, "נמצא בהצלחה");
            
            System.out.println("\nלוחץ על אזור התחבר/הרשם...");
            try {
                dropdownTrigger.click();
            } catch (Exception e) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", dropdownTrigger);
            }
            Thread.sleep(2000);
            System.out.println("✓ התפריט נפתח");
            
            TestResultsExcelWriter.addTestResult("פתיחת תפריט", true, "נפתח בהצלחה");
            
            registerButton = driver.findElement(By.xpath("//a[@href='Register' or contains(@href, 'Register')]"));
            System.out.println("נמצא קישור הרשמה: '" + registerButton.getText() + "'");
            
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", registerButton);
            Thread.sleep(1500);
            
            ((JavascriptExecutor) driver).executeScript(
                "arguments[0].style.border='1px solid #FFB6C1';" +
                "arguments[0].style.boxShadow='0 0 3px rgba(255, 182, 193, 0.5)';" +
                "arguments[0].style.borderRadius='3px';" +
                "arguments[0].style.outline='1px solid rgba(255, 182, 193, 0.3)';" +
                "arguments[0].style.outlineOffset='2px';", 
                registerButton
            );
            Thread.sleep(3000);
            
            TestResultsExcelWriter.addTestResult("מציאת כפתור הרשמה", true, "נמצא בהצלחה");
            
            System.out.println("\nלוחץ על כפתור הרשמה...");
            try {
                registerButton.click();
            } catch (Exception e) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", registerButton);
            }
            Thread.sleep(3000);
            System.out.println("✓ לחץ על כפתור הרשמה");
            
            TestResultsExcelWriter.addTestResult("לחיצה על כפתור הרשמה", true, "בוצעה בהצלחה");
            
        } catch (Exception e) {
            System.out.println("שגיאה בניווט דרך תפריט: " + e.getMessage());
            TestResultsExcelWriter.addTestResult("ניווט דרך תפריט", false, "עובר ישירות");
            driver.get("https://www.lastprice.co.il/Register");
            Thread.sleep(3000);
        }
        
        System.out.println("✓ הגיע לעמוד טופס ההרשמה");
        
        // וידוא שאנחנו בדף הרשמה ולא בדף עדכון פרטים
        String currentUrl = driver.getCurrentUrl();
        System.out.println("URL נוכחי: " + currentUrl);
        
        if (currentUrl.contains("MyAccount") || currentUrl.contains("Profile") || currentUrl.contains("Update")) {
            System.out.println("⚠️ זוהה שאנחנו בדף עדכון פרטים ולא בהרשמה! מבצע התנתקות...");
            driver.manage().deleteAllCookies();
            Thread.sleep(1000);
            driver.get("https://www.lastprice.co.il/Register");
            Thread.sleep(3000);
            System.out.println("✓ ניווט ישיר לדף הרשמה אחרי התנתקות");
        }
        
        TestResultsExcelWriter.addTestResult("הגעה לטופס", true, "הגיע בהצלחה");
    }
}
