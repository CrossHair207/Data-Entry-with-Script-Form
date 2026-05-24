SOURCE link: https://www.youtube.com/watch?v=_u-ubZVDF7M


-----------------
Deployment successfully updated.
Version 1 on May 24, 2026, 8:00 PM
Deployment ID: AKfycbzxgsi0impmo3RnZP8Iies1FcViJrdaHiXtAm9IAbhBoNTSMxOVtYfa2KcfBRSEfpB4Hg

Web app
URL: https://script.google.com/macros/s/AKfycbzxgsi0impmo3RnZP8Iies1FcViJrdaHiXtAm9IAbhBoNTSMxOVtYfa2KcfBRSEfpB4Hg/exec
------------------

## Setup (Google Apps Script)

1. Open the spreadsheet **Data Entry with Script Form**
2. Go to **Extensions → Apps Script**
3. Paste **Code.gs** into `Code.gs`
4. Add an HTML file named **index** and paste **index.html** into it
5. Save the project
6. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (or **Anyone with Google account**)
7. Open the **Web app URL** above (do not open `index.html` as a local file)

## "This app hasn't been verified by Google" — how to continue

This warning is normal for personal Apps Script projects. Your code is fine.

1. Click **Advanced** (bottom left)
2. Click **Go to [your app name] (unsafe)**
3. Choose your Google account
4. Click **Allow**

Other users (VA / team) can use the same steps if they trust you and your script.

## Files in this project

| File        | Purpose                                      |
|-------------|----------------------------------------------|
| Code.gs     | Backend — reads/writes Google Sheets         |
| index.html  | Data entry form UI (use as Apps Script `index`) |

## Spreadsheet columns (row 1)

First Name, Last Name, Full Name, Email Address, PhoneNumber, Mobile Number, Date Of Birth, Gender, Company Name, Job Title, Department, Website URL, LinkedIn Profile, Street Address, City, State / Province, ZIP / Postal Code, Country, Customer ID, Employee ID, Order Number, Invoice Number, Database Entry ID, Product Name, Product SKU, Quantity, Price, Payment Status, Shipping Status, Appointment Date, Meeting Time, Account Username, Registration Date, Subscription Status, Tags / Categories, Social Media Links, Notes / Comments

The script creates or fixes these headers automatically if they are missing.

## Troubleshooting

- **Failed to load records** — Open the form only via the deployed Web app URL, not by double-clicking `index.html` on your PC.
- **Spreadsheet not found** — The spreadsheet name must be exactly **Data Entry with Script Form**.
- **Save failed** — First Name, Last Name, and a valid Email are required.
- **After code changes** — Deploy → **Manage deployments** → Edit → **New version** → Deploy.
