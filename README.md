# BISSCOTI Manager
מערכת ניהול מרכזית למפעל Biscotti

## מבנה הפרויקט

```
BISSCOTI-manager/
├── Code.gs/          ← קוד שרת Google Apps Script
│   └── Code.gs       ← נתב ראשי (doGet / doPost)
└── html/             ← קוד Frontend (Netlify)
    └── אחזקה.html    ← מסך ניהול קריאות אחזקה
```

## Code.gs – Actions נתמכים

| Method | action        | תיאור                    |
|--------|---------------|--------------------------|
| GET    | login         | התחברות משתמש            |
| GET    | getUsers      | רשימת משתמשים (מנהל)    |
| GET    | —             | טעינת כל הטיקטים         |
| POST   | ADD           | הוספת טיקט               |
| POST   | UPDATE        | עדכון טיקט               |
| POST   | DELETE        | מחיקת טיקט               |
| POST   | addUser       | הוספת משתמש              |
| POST   | deleteUser    | מחיקת משתמש              |

## התחלה מהירה

1. פתח את `Code.gs/Code.gs` ב-Apps Script
2. הגדר `SHEET_ID` למזהה ה-Google Sheet
3. ודא שיש שני Sheets: `tickets` ו-`users` עם headers תואמים
4. פרוס כ-Web App ועדכן `API_URL` ב-html
